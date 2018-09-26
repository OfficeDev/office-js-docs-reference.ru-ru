
# <a name="item"></a><span data-ttu-id="18cff-101">item</span><span class="sxs-lookup"><span data-stu-id="18cff-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="18cff-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="18cff-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="18cff-p101">Пространство имен `item` используется для доступа к выбранному в данный момент сообщению, приглашению на собрание или описанию встречи. Вы можете определить тип пространства имен `item` с помощью свойства [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="18cff-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="18cff-105">Requirements</span></span>

|<span data-ttu-id="18cff-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-106">Requirement</span></span>|<span data-ttu-id="18cff-107">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-109">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-109">1.0</span></span>|
|[<span data-ttu-id="18cff-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="18cff-111">Restricted</span></span>|
|[<span data-ttu-id="18cff-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="18cff-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="18cff-114">Members and methods</span></span>

| <span data-ttu-id="18cff-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-115">Member</span></span> | <span data-ttu-id="18cff-116">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="18cff-117">attachments</span><span class="sxs-lookup"><span data-stu-id="18cff-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="18cff-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-118">Member</span></span> |
| [<span data-ttu-id="18cff-119">bcc</span><span class="sxs-lookup"><span data-stu-id="18cff-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="18cff-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-120">Member</span></span> |
| [<span data-ttu-id="18cff-121">body</span><span class="sxs-lookup"><span data-stu-id="18cff-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="18cff-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-122">Member</span></span> |
| [<span data-ttu-id="18cff-123">cc</span><span class="sxs-lookup"><span data-stu-id="18cff-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="18cff-124">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-124">Member</span></span> |
| [<span data-ttu-id="18cff-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="18cff-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="18cff-126">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-126">Member</span></span> |
| [<span data-ttu-id="18cff-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="18cff-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="18cff-128">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-128">Member</span></span> |
| [<span data-ttu-id="18cff-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="18cff-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="18cff-130">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-130">Member</span></span> |
| [<span data-ttu-id="18cff-131">end</span><span class="sxs-lookup"><span data-stu-id="18cff-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="18cff-132">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-132">Member</span></span> |
| [<span data-ttu-id="18cff-133">from</span><span class="sxs-lookup"><span data-stu-id="18cff-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="18cff-134">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-134">Member</span></span> |
| [<span data-ttu-id="18cff-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="18cff-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="18cff-136">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-136">Member</span></span> |
| [<span data-ttu-id="18cff-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="18cff-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="18cff-138">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-138">Member</span></span> |
| [<span data-ttu-id="18cff-139">itemId</span><span class="sxs-lookup"><span data-stu-id="18cff-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="18cff-140">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-140">Member</span></span> |
| [<span data-ttu-id="18cff-141">itemType</span><span class="sxs-lookup"><span data-stu-id="18cff-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="18cff-142">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-142">Member</span></span> |
| [<span data-ttu-id="18cff-143">location</span><span class="sxs-lookup"><span data-stu-id="18cff-143">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="18cff-144">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-144">Member</span></span> |
| [<span data-ttu-id="18cff-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="18cff-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="18cff-146">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-146">Member</span></span> |
| [<span data-ttu-id="18cff-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="18cff-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="18cff-148">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-148">Member</span></span> |
| [<span data-ttu-id="18cff-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="18cff-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="18cff-150">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-150">Member</span></span> |
| [<span data-ttu-id="18cff-151">organizer</span><span class="sxs-lookup"><span data-stu-id="18cff-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="18cff-152">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-152">Member</span></span> |
| [<span data-ttu-id="18cff-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="18cff-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="18cff-154">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-154">Member</span></span> |
| [<span data-ttu-id="18cff-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="18cff-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="18cff-156">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-156">Member</span></span> |
| [<span data-ttu-id="18cff-157">sender</span><span class="sxs-lookup"><span data-stu-id="18cff-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="18cff-158">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-158">Member</span></span> |
| [<span data-ttu-id="18cff-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="18cff-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="18cff-160">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-160">Member</span></span> |
| [<span data-ttu-id="18cff-161">start</span><span class="sxs-lookup"><span data-stu-id="18cff-161">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="18cff-162">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-162">Member</span></span> |
| [<span data-ttu-id="18cff-163">subject</span><span class="sxs-lookup"><span data-stu-id="18cff-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="18cff-164">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-164">Member</span></span> |
| [<span data-ttu-id="18cff-165">to</span><span class="sxs-lookup"><span data-stu-id="18cff-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="18cff-166">Элемент</span><span class="sxs-lookup"><span data-stu-id="18cff-166">Member</span></span> |
| [<span data-ttu-id="18cff-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="18cff-168">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-168">Method</span></span> |
| [<span data-ttu-id="18cff-169">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="18cff-169">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="18cff-170">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-170">Method</span></span> |
| [<span data-ttu-id="18cff-171">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-171">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="18cff-172">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-172">Method</span></span> |
| [<span data-ttu-id="18cff-173">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-173">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="18cff-174">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-174">Method</span></span> |
| [<span data-ttu-id="18cff-175">close</span><span class="sxs-lookup"><span data-stu-id="18cff-175">close</span></span>](#close) | <span data-ttu-id="18cff-176">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-176">Method</span></span> |
| [<span data-ttu-id="18cff-177">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="18cff-177">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="18cff-178">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-178">Method</span></span> |
| [<span data-ttu-id="18cff-179">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="18cff-179">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="18cff-180">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-180">Method</span></span> |
| [<span data-ttu-id="18cff-181">getEntities</span><span class="sxs-lookup"><span data-stu-id="18cff-181">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="18cff-182">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-182">Method</span></span> |
| [<span data-ttu-id="18cff-183">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="18cff-183">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="18cff-184">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-184">Method</span></span> |
| [<span data-ttu-id="18cff-185">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="18cff-185">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="18cff-186">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-186">Method</span></span> |
| [<span data-ttu-id="18cff-187">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-187">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="18cff-188">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-188">Method</span></span> |
| [<span data-ttu-id="18cff-189">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="18cff-189">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="18cff-190">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-190">Method</span></span> |
| [<span data-ttu-id="18cff-191">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="18cff-191">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="18cff-192">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-192">Method</span></span> |
| [<span data-ttu-id="18cff-193">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-193">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="18cff-194">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-194">Method</span></span> |
| [<span data-ttu-id="18cff-195">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="18cff-195">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="18cff-196">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-196">Method</span></span> |
| [<span data-ttu-id="18cff-197">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="18cff-197">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="18cff-198">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-198">Method</span></span> |
| [<span data-ttu-id="18cff-199">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-199">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="18cff-200">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-200">Method</span></span> |
| [<span data-ttu-id="18cff-201">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-201">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="18cff-202">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-202">Method</span></span> |
| [<span data-ttu-id="18cff-203">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-203">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="18cff-204">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-204">Method</span></span> |
| [<span data-ttu-id="18cff-205">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-205">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="18cff-206">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-206">Method</span></span> |
| [<span data-ttu-id="18cff-207">saveAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-207">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="18cff-208">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-208">Method</span></span> |
| [<span data-ttu-id="18cff-209">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="18cff-209">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="18cff-210">Метод</span><span class="sxs-lookup"><span data-stu-id="18cff-210">Method</span></span> |

### <a name="example"></a><span data-ttu-id="18cff-211">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-211">Example</span></span>

<span data-ttu-id="18cff-212">В примере кода JavaScript, приведенном ниже, показано, как получить доступ к свойству `subject` текущего элемента в Outlook.</span><span class="sxs-lookup"><span data-stu-id="18cff-212">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="18cff-213">Элементы</span><span class="sxs-lookup"><span data-stu-id="18cff-213">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="18cff-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="18cff-214">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="18cff-p102">Получает массив вложений для элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-217">Определенные типы файлов блокируемых в Outlook из-за потенциальных проблем безопасности и поэтому не возвращаются.</span><span class="sxs-lookup"><span data-stu-id="18cff-217">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="18cff-218">Для получения дополнительных сведений см [Блокировка вложений в Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="18cff-218">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-219">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-219">Type:</span></span>

*   <span data-ttu-id="18cff-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="18cff-220">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-221">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-221">Requirements</span></span>

|<span data-ttu-id="18cff-222">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-222">Requirement</span></span>|<span data-ttu-id="18cff-223">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-225">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-225">1.0</span></span>|
|[<span data-ttu-id="18cff-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-227">ReadItem</span></span>|
|[<span data-ttu-id="18cff-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-229">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-230">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-230">Example</span></span>

<span data-ttu-id="18cff-231">С помощью приведенного ниже кода можно создать HTML-строку с подробными сведениями обо всех вложениях для текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-231">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="18cff-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-232">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="18cff-233">Получает объект, который предоставляет методы для получения или обновления получателей в строке (Скрытая копия) скрытой копии сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-233">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="18cff-234">Только в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-234">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-235">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-235">Type:</span></span>

*   [<span data-ttu-id="18cff-236">Recipients</span><span class="sxs-lookup"><span data-stu-id="18cff-236">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="18cff-237">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-237">Requirements</span></span>

|<span data-ttu-id="18cff-238">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-238">Requirement</span></span>|<span data-ttu-id="18cff-239">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-240">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-241">1.1</span><span class="sxs-lookup"><span data-stu-id="18cff-241">1.1</span></span>|
|[<span data-ttu-id="18cff-242">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-243">ReadItem</span></span>|
|[<span data-ttu-id="18cff-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-245">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-245">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-246">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-246">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="18cff-247">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="18cff-247">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="18cff-248">Получает объект, предоставляющий методы для работы с основным текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-248">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-249">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-249">Type:</span></span>

*   [<span data-ttu-id="18cff-250">Body</span><span class="sxs-lookup"><span data-stu-id="18cff-250">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="18cff-251">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-251">Requirements</span></span>

|<span data-ttu-id="18cff-252">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-252">Requirement</span></span>|<span data-ttu-id="18cff-253">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-254">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-254">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-255">1.1</span><span class="sxs-lookup"><span data-stu-id="18cff-255">1.1</span></span>|
|[<span data-ttu-id="18cff-256">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-256">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-257">ReadItem</span></span>|
|[<span data-ttu-id="18cff-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-258">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-259">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="18cff-260">cc: массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получателей](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-260">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="18cff-261">Предоставляет доступ к «копия» (копия) получателей сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-261">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="18cff-262">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-262">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-263">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-263">Read mode</span></span>

<span data-ttu-id="18cff-p106">Свойство `cc` возвращает массив, который содержит объект `EmailAddressDetails` для каждого получателя, указанного в строке **Копия** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-266">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-266">Compose mode</span></span>

<span data-ttu-id="18cff-267">`cc` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **копия** сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-267">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-268">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-268">Type:</span></span>

*   <span data-ttu-id="18cff-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-269">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-270">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-270">Requirements</span></span>

|<span data-ttu-id="18cff-271">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-271">Requirement</span></span>|<span data-ttu-id="18cff-272">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-274">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-274">1.0</span></span>|
|[<span data-ttu-id="18cff-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-276">ReadItem</span></span>|
|[<span data-ttu-id="18cff-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-278">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-278">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-279">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-279">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="18cff-280">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="18cff-280">(nullable) conversationId :String</span></span>

<span data-ttu-id="18cff-281">Получает идентификатор разговора по электронной почте, содержащего конкретное сообщение.</span><span class="sxs-lookup"><span data-stu-id="18cff-281">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="18cff-p107">Вы можете получить целочисленное значение этого свойства, если ваше почтовое приложение активируется в формах просмотра или формах создания ответов. Если пользователь изменит тему ответа, после его отправки идентификатор беседы будет изменен, и полученное ранее значение будет недействительным.</span><span class="sxs-lookup"><span data-stu-id="18cff-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="18cff-p108">Это свойство имеет значение NULL для нового элемента в форме создания. Свойство `conversationId` вернет значение, если пользователь задаст тему и сохранит элемент.</span><span class="sxs-lookup"><span data-stu-id="18cff-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-286">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-286">Type:</span></span>

*   <span data-ttu-id="18cff-287">String</span><span class="sxs-lookup"><span data-stu-id="18cff-287">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-288">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-288">Requirements</span></span>

|<span data-ttu-id="18cff-289">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-289">Requirement</span></span>|<span data-ttu-id="18cff-290">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-291">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-292">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-292">1.0</span></span>|
|[<span data-ttu-id="18cff-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-294">ReadItem</span></span>|
|[<span data-ttu-id="18cff-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-296">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-296">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="18cff-297">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="18cff-297">dateTimeCreated :Date</span></span>

<span data-ttu-id="18cff-p109">Получает дату и время создания элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-300">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-300">Type:</span></span>

*   <span data-ttu-id="18cff-301">Date</span><span class="sxs-lookup"><span data-stu-id="18cff-301">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-302">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-302">Requirements</span></span>

|<span data-ttu-id="18cff-303">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-303">Requirement</span></span>|<span data-ttu-id="18cff-304">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-305">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-305">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-306">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-306">1.0</span></span>|
|[<span data-ttu-id="18cff-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-308">ReadItem</span></span>|
|[<span data-ttu-id="18cff-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-310">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-311">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-311">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="18cff-312">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="18cff-312">dateTimeModified :Date</span></span>

<span data-ttu-id="18cff-p110">Получает дату и время последнего изменения элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-315">Этот член не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-315">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-316">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-316">Type:</span></span>

*   <span data-ttu-id="18cff-317">Date</span><span class="sxs-lookup"><span data-stu-id="18cff-317">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-318">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-318">Requirements</span></span>

|<span data-ttu-id="18cff-319">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-319">Requirement</span></span>|<span data-ttu-id="18cff-320">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-320">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-321">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-321">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-322">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-322">1.0</span></span>|
|[<span data-ttu-id="18cff-323">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-323">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-324">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-324">ReadItem</span></span>|
|[<span data-ttu-id="18cff-325">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-325">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-326">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-326">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-327">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-327">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="18cff-328">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="18cff-328">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="18cff-329">Получает или задает дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-329">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="18cff-p111">Свойство `end` представлено в виде значения даты и времени в формате UTC. Преобразовать значение свойства end в местные значения даты и времени клиента можно с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="18cff-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-332">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-332">Read mode</span></span>

<span data-ttu-id="18cff-333">Свойство `end` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="18cff-333">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-334">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-334">Compose mode</span></span>

<span data-ttu-id="18cff-335">Свойство `end` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="18cff-335">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="18cff-336">Если вы задаете время окончания с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="18cff-336">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-337">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-337">Type:</span></span>

*   <span data-ttu-id="18cff-338">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="18cff-338">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-339">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-339">Requirements</span></span>

|<span data-ttu-id="18cff-340">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-340">Requirement</span></span>|<span data-ttu-id="18cff-341">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-342">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-342">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-343">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-343">1.0</span></span>|
|[<span data-ttu-id="18cff-344">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-345">ReadItem</span></span>|
|[<span data-ttu-id="18cff-346">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-347">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-348">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-348">Example</span></span>

<span data-ttu-id="18cff-349">В примере ниже показано, как с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задать время окончания встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-349">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="18cff-350">от:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="18cff-350">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="18cff-351">Получает адрес электронной почты отправителя сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-351">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="18cff-p112">Свойства `from` и [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="18cff-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-354">`recipientType` Свойства `EmailAddressDetails` объект в `from` — это свойство `undefined`.</span><span class="sxs-lookup"><span data-stu-id="18cff-354">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-355">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-355">Read mode</span></span>

<span data-ttu-id="18cff-356">`from` Возвращает свойство `EmailAddressDetails` объекта.</span><span class="sxs-lookup"><span data-stu-id="18cff-356">The `from` property returns an `EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="18cff-357">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-357">Compose mode</span></span>

<span data-ttu-id="18cff-358">`from` Возвращает свойство `From` объект, который предоставляет метод для получения из значения.</span><span class="sxs-lookup"><span data-stu-id="18cff-358">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="18cff-359">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-359">Type:</span></span>

*   <span data-ttu-id="18cff-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [из](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="18cff-360">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-361">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-361">Requirements</span></span>

|<span data-ttu-id="18cff-362">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-362">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="18cff-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-363">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-364">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-364">1.0</span></span>|<span data-ttu-id="18cff-365">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-365">1.7</span></span>|
|[<span data-ttu-id="18cff-366">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-366">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-367">ReadItem</span></span>|<span data-ttu-id="18cff-368">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-368">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-369">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-369">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-370">Read</span><span class="sxs-lookup"><span data-stu-id="18cff-370">Read</span></span>|<span data-ttu-id="18cff-371">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-371">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="18cff-372">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="18cff-372">internetMessageId :String</span></span>

<span data-ttu-id="18cff-p113">Получает идентификатор интернет-сообщения для электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-375">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-375">Type:</span></span>

*   <span data-ttu-id="18cff-376">String</span><span class="sxs-lookup"><span data-stu-id="18cff-376">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-377">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-377">Requirements</span></span>

|<span data-ttu-id="18cff-378">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-378">Requirement</span></span>|<span data-ttu-id="18cff-379">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-380">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-380">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-381">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-381">1.0</span></span>|
|[<span data-ttu-id="18cff-382">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-382">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-383">ReadItem</span></span>|
|[<span data-ttu-id="18cff-384">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-384">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-385">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-385">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-386">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-386">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="18cff-387">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="18cff-387">itemClass :String</span></span>

<span data-ttu-id="18cff-p114">Получает класс элемента веб-служб Exchange для выбранного элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="18cff-p115">Свойство `itemClass` указывает класс сообщения выбранного элемента. Ниже приводятся классы сообщения по умолчанию для элемента сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="18cff-392">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-392">Type</span></span>|<span data-ttu-id="18cff-393">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-393">Description</span></span>|<span data-ttu-id="18cff-394">Класс элемента</span><span class="sxs-lookup"><span data-stu-id="18cff-394">item class</span></span>|
|---|---|---|
|<span data-ttu-id="18cff-395">Элементы встречи</span><span class="sxs-lookup"><span data-stu-id="18cff-395">Appointment items</span></span>|<span data-ttu-id="18cff-396">Это элементы календаря для класса элемента `IPM.Appointment` или `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="18cff-396">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="18cff-397">Элементы сообщения</span><span class="sxs-lookup"><span data-stu-id="18cff-397">Message items</span></span>|<span data-ttu-id="18cff-398">Сюда входят электронные сообщения, для которых по умолчанию задан класс сообщения `IPM.Note`, а также приглашения на собрания, ответы на них и уведомления об их отмене, использующие `IPM.Schedule.Meeting` в качестве базового класса сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-398">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="18cff-399">Можно создавать настраиваемые классы сообщения, расширяющие классы сообщения по умолчанию, например настраиваемый класс сообщения о встрече `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="18cff-399">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-400">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-400">Type:</span></span>

*   <span data-ttu-id="18cff-401">String</span><span class="sxs-lookup"><span data-stu-id="18cff-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-402">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-402">Requirements</span></span>

|<span data-ttu-id="18cff-403">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-403">Requirement</span></span>|<span data-ttu-id="18cff-404">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-405">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-405">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-406">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-406">1.0</span></span>|
|[<span data-ttu-id="18cff-407">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-407">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-408">ReadItem</span></span>|
|[<span data-ttu-id="18cff-409">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-409">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-410">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-411">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-411">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="18cff-412">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="18cff-412">(nullable) itemId :String</span></span>

<span data-ttu-id="18cff-p116">Получает идентификатор элемента веб-служб Exchange для текущего элемента. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-415">Идентификатор, возвращаемый свойством `itemId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="18cff-415">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="18cff-416">`itemId` Свойство не совпадать с Идентификатором, используемым API-Интерфейс REST Outlook или идентификатор записи Outlook.</span><span class="sxs-lookup"><span data-stu-id="18cff-416">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="18cff-417">Прежде чем вносить API-Интерфейс REST для звонков с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="18cff-417">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="18cff-418">Для получения дополнительных сведений показано [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="18cff-418">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="18cff-p118">Свойство `itemId` недоступно в режиме создания. Если требуется идентификатор элемента, с помощью метода [`saveAsync`](#saveasyncoptions-callback) можно сохранить элемент в хранилище. При этом в параметре [`AsyncResult.value`](/javascript/api/office/office.asyncresult) функции обратного вызова возвращается идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-421">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-421">Type:</span></span>

*   <span data-ttu-id="18cff-422">String</span><span class="sxs-lookup"><span data-stu-id="18cff-422">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-423">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-423">Requirements</span></span>

|<span data-ttu-id="18cff-424">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-424">Requirement</span></span>|<span data-ttu-id="18cff-425">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-426">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-427">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-427">1.0</span></span>|
|[<span data-ttu-id="18cff-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-428">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-429">ReadItem</span></span>|
|[<span data-ttu-id="18cff-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-430">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-431">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-432">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-432">Example</span></span>

<span data-ttu-id="18cff-p119">Указанный ниже код проверяет наличие идентификатора элемента. Если свойство `itemId` возвращает значение `null` или `undefined`, элемент будет сохранен в хранилище, а из асинхронного результата будет получен идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="18cff-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="18cff-435">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="18cff-436">Получает тип элемента, который представляет экземпляр.</span><span class="sxs-lookup"><span data-stu-id="18cff-436">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="18cff-437">Свойство `itemType` возвращает одно из значений перечисления `ItemType`, которое указывает, является ли экземпляр объекта `item` сообщением или встречей.</span><span class="sxs-lookup"><span data-stu-id="18cff-437">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-438">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-438">Type:</span></span>

*   [<span data-ttu-id="18cff-439">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="18cff-439">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="18cff-440">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-440">Requirements</span></span>

|<span data-ttu-id="18cff-441">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-441">Requirement</span></span>|<span data-ttu-id="18cff-442">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-443">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-443">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-444">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-444">1.0</span></span>|
|[<span data-ttu-id="18cff-445">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-446">ReadItem</span></span>|
|[<span data-ttu-id="18cff-447">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-448">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-449">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-449">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="18cff-450">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="18cff-450">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="18cff-451">Получает или задает место встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-451">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-452">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-452">Read mode</span></span>

<span data-ttu-id="18cff-453">Свойство `location` возвращает строку, содержащую сведения о месте встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-453">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-454">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-454">Compose mode</span></span>

<span data-ttu-id="18cff-455">Свойство `location` возвращает объект `Location`, предоставляющий методы, которые используются для получения и задания места встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-455">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-456">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-456">Type:</span></span>

*   <span data-ttu-id="18cff-457">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="18cff-457">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-458">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-458">Requirements</span></span>

|<span data-ttu-id="18cff-459">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-459">Requirement</span></span>|<span data-ttu-id="18cff-460">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-461">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-461">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-462">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-462">1.0</span></span>|
|[<span data-ttu-id="18cff-463">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-464">ReadItem</span></span>|
|[<span data-ttu-id="18cff-465">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-466">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-466">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-467">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-467">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="18cff-468">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="18cff-468">normalizedSubject :String</span></span>

<span data-ttu-id="18cff-p120">Получает тему элемента со всеми удаленными префиксами (включая `RE:` и `FWD:`). Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="18cff-p121">Свойство normalizedSubject получает тему элемента со стандартными префиксами (такими как `RE:` и `FW:`), добавляемыми почтовыми программами. Для получения темы элемента с неизмененными префиксами используйте свойство [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="18cff-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-473">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-473">Type:</span></span>

*   <span data-ttu-id="18cff-474">String</span><span class="sxs-lookup"><span data-stu-id="18cff-474">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-475">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-475">Requirements</span></span>

|<span data-ttu-id="18cff-476">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-476">Requirement</span></span>|<span data-ttu-id="18cff-477">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-477">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-478">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-478">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-479">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-479">1.0</span></span>|
|[<span data-ttu-id="18cff-480">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-480">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-481">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-481">ReadItem</span></span>|
|[<span data-ttu-id="18cff-482">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-482">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-483">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-483">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-484">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-484">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="18cff-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="18cff-485">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="18cff-486">Получает сообщения уведомления для элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-486">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-487">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-487">Type:</span></span>

*   [<span data-ttu-id="18cff-488">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="18cff-488">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="18cff-489">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-489">Requirements</span></span>

|<span data-ttu-id="18cff-490">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-490">Requirement</span></span>|<span data-ttu-id="18cff-491">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-492">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-492">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-493">1.3</span><span class="sxs-lookup"><span data-stu-id="18cff-493">1.3</span></span>|
|[<span data-ttu-id="18cff-494">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-495">ReadItem</span></span>|
|[<span data-ttu-id="18cff-496">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-497">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-497">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="18cff-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-498">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="18cff-499">Предоставляет доступ к необязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="18cff-499">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="18cff-500">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-500">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-501">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-501">Read mode</span></span>

<span data-ttu-id="18cff-502">Свойство `optionalAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого необязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-502">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-503">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-503">Compose mode</span></span>

<span data-ttu-id="18cff-504">`optionalAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления необязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-504">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-505">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-505">Type:</span></span>

*   <span data-ttu-id="18cff-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-507">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-507">Requirements</span></span>

|<span data-ttu-id="18cff-508">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-508">Requirement</span></span>|<span data-ttu-id="18cff-509">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-510">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-510">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-511">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-511">1.0</span></span>|
|[<span data-ttu-id="18cff-512">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-512">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-513">ReadItem</span></span>|
|[<span data-ttu-id="18cff-514">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-514">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-515">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-515">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-516">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-516">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="18cff-517">Организатор:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[организатора](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="18cff-517">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="18cff-518">Получает адрес электронной почты организатора указанного собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-518">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-519">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-519">Read mode</span></span>

<span data-ttu-id="18cff-520">`organizer` Свойство возвращает объект [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) , который представляет организатором собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-520">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-521">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-521">Compose mode</span></span>

<span data-ttu-id="18cff-522">`organizer` Свойство возвращает объект [организатора](/javascript/api/outlook/office.organizer) , который предоставляет метод для получения значения Организатор.</span><span class="sxs-lookup"><span data-stu-id="18cff-522">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-523">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-523">Type:</span></span>

*   <span data-ttu-id="18cff-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [организатора](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="18cff-524">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-525">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-525">Requirements</span></span>

|<span data-ttu-id="18cff-526">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-526">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="18cff-527">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-527">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-528">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-528">1.0</span></span>|<span data-ttu-id="18cff-529">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-529">1.7</span></span>|
|[<span data-ttu-id="18cff-530">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-530">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-531">ReadItem</span></span>|<span data-ttu-id="18cff-532">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-532">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-533">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-534">Read</span><span class="sxs-lookup"><span data-stu-id="18cff-534">Read</span></span>|<span data-ttu-id="18cff-535">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-536">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-536">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="18cff-537">(значение null) повторения:[повторения](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="18cff-537">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="18cff-538">Получает или задает шаблон повторения встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-538">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="18cff-539">Получает шаблон повторения приглашения на собрание.</span><span class="sxs-lookup"><span data-stu-id="18cff-539">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="18cff-540">Читать и создавать режимы для элементов встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-540">Read and compose modes for appointment items.</span></span> <span data-ttu-id="18cff-541">В режиме чтения к собранию элементы запроса.</span><span class="sxs-lookup"><span data-stu-id="18cff-541">Read mode for meeting request items.</span></span>

<span data-ttu-id="18cff-542">`recurrence` При элемента ряд или экземпляра в цикле свойство возвращает объект [повторения](/javascript/api/outlook/office.recurrence) для повторяющиеся встречи или собрания запросы.</span><span class="sxs-lookup"><span data-stu-id="18cff-542">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="18cff-543">`null`возвращаются для одного встреч и приглашений на собрания из одного встреч.</span><span class="sxs-lookup"><span data-stu-id="18cff-543">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="18cff-544">`undefined`возвращается для сообщений, которые не являются приглашений на собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-544">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="18cff-545">Примечание: Приглашений на собрание имеют `itemClass` значение IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="18cff-545">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="18cff-546">Примечание: Если объект повторения `null`, это означает, что объект является одной встречи или приглашения на собрание из одной встречи и не является частью серии.</span><span class="sxs-lookup"><span data-stu-id="18cff-546">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-547">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-547">Type:</span></span>

* [<span data-ttu-id="18cff-548">Повторение</span><span class="sxs-lookup"><span data-stu-id="18cff-548">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="18cff-549">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-549">Requirement</span></span>|<span data-ttu-id="18cff-550">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-551">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-551">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-552">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-552">1.7</span></span>|
|[<span data-ttu-id="18cff-553">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-554">ReadItem</span></span>|
|[<span data-ttu-id="18cff-555">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-556">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-556">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="18cff-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-557">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="18cff-558">Предоставляет доступ к обязательным участникам события.</span><span class="sxs-lookup"><span data-stu-id="18cff-558">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="18cff-559">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-559">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-560">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-560">Read mode</span></span>

<span data-ttu-id="18cff-561">Свойство `requiredAttendees` возвращает массив, содержащий объект `EmailAddressDetails` для каждого обязательного участника собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-561">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-562">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-562">Compose mode</span></span>

<span data-ttu-id="18cff-563">`requiredAttendees` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления обязательных участников собрания.</span><span class="sxs-lookup"><span data-stu-id="18cff-563">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-564">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-564">Type:</span></span>

*   <span data-ttu-id="18cff-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-565">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-566">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-566">Requirements</span></span>

|<span data-ttu-id="18cff-567">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-567">Requirement</span></span>|<span data-ttu-id="18cff-568">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-568">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-569">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-569">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-570">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-570">1.0</span></span>|
|[<span data-ttu-id="18cff-571">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-572">ReadItem</span></span>|
|[<span data-ttu-id="18cff-573">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-573">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-574">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-574">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-575">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-575">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="18cff-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="18cff-576">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="18cff-p126">Получает электронный адрес отправителя электронного сообщения. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="18cff-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="18cff-p127">Свойства [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) и `sender` представляют одно лицо, если сообщение не отправлено представителем. В противном случае свойство `from` представляет лицо, делегировавшее полномочия, а свойство sender — представителя.</span><span class="sxs-lookup"><span data-stu-id="18cff-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-581">`recipientType` Свойства `EmailAddressDetails` объект в `sender` — это свойство `undefined`.</span><span class="sxs-lookup"><span data-stu-id="18cff-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-582">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-582">Type:</span></span>

*   [<span data-ttu-id="18cff-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="18cff-583">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="18cff-584">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-584">Requirements</span></span>

|<span data-ttu-id="18cff-585">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-585">Requirement</span></span>|<span data-ttu-id="18cff-586">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-587">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-587">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-588">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-588">1.0</span></span>|
|[<span data-ttu-id="18cff-589">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-590">ReadItem</span></span>|
|[<span data-ttu-id="18cff-591">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-592">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-593">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-593">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="18cff-594">(значение null) seriesId: String</span><span class="sxs-lookup"><span data-stu-id="18cff-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="18cff-595">Получает идентификатор серии, к которой принадлежит экземпляр.</span><span class="sxs-lookup"><span data-stu-id="18cff-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="18cff-596">В OWA и Outlook `seriesId` возвращает идентификатор веб-служб Exchange (EWS) элемента родительского (ряды), к которому принадлежит этот элемент.</span><span class="sxs-lookup"><span data-stu-id="18cff-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="18cff-597">Однако в iOS и Android `seriesId` возвращает REST идентификатор родительского элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-598">Идентификатор, возвращаемый свойством `seriesId`, совпадает с идентификатором элемента веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="18cff-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="18cff-599">`seriesId` Свойство не идентичен идентификаторы Outlook, используемые API-Интерфейс REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="18cff-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="18cff-600">Прежде чем вносить API-Интерфейс REST для звонков с помощью этого значения, его следует преобразовать с помощью [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="18cff-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="18cff-601">Для получения дополнительных сведений показано [Использование API REST Outlook из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="18cff-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="18cff-602">`seriesId` Возвращает свойство `null` для элементов, не имеющих родительских элементов, таких как единый встреч, элементы ряда или собрания запрашивает и возвращает `undefined` для других элементов, которые не являются соответствующие запросы.</span><span class="sxs-lookup"><span data-stu-id="18cff-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-603">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-603">Type:</span></span>

* <span data-ttu-id="18cff-604">String</span><span class="sxs-lookup"><span data-stu-id="18cff-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-605">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-605">Requirements</span></span>

|<span data-ttu-id="18cff-606">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-606">Requirement</span></span>|<span data-ttu-id="18cff-607">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-608">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-608">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-609">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-609">1.7</span></span>|
|[<span data-ttu-id="18cff-610">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-611">ReadItem</span></span>|
|[<span data-ttu-id="18cff-612">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-613">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-613">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-614">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-614">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="18cff-615">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="18cff-615">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="18cff-616">Получает или задает дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="18cff-p130">Свойство `start` представлено в виде значения даты и времени в формате UTC. Это значение можно преобразовать в местные значения даты и времени клиента с помощью метода [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime).</span><span class="sxs-lookup"><span data-stu-id="18cff-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-619">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-619">Read mode</span></span>

<span data-ttu-id="18cff-620">Свойство `start` возвращает объект `Date`.</span><span class="sxs-lookup"><span data-stu-id="18cff-620">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-621">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-621">Compose mode</span></span>

<span data-ttu-id="18cff-622">Свойство `start` возвращает объект `Time`.</span><span class="sxs-lookup"><span data-stu-id="18cff-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="18cff-623">Если вы задаете время начала с помощью метода [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-), необходимо использовать метод [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) для преобразования местного времени на клиенте в формат UTC для сервера.</span><span class="sxs-lookup"><span data-stu-id="18cff-623">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-624">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-624">Type:</span></span>

*   <span data-ttu-id="18cff-625">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="18cff-625">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-626">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-626">Requirements</span></span>

|<span data-ttu-id="18cff-627">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-627">Requirement</span></span>|<span data-ttu-id="18cff-628">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-629">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-629">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-630">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-630">1.0</span></span>|
|[<span data-ttu-id="18cff-631">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-632">ReadItem</span></span>|
|[<span data-ttu-id="18cff-633">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-634">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-634">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-635">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-635">Example</span></span>

<span data-ttu-id="18cff-636">В примере ниже с помощью метода [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) объекта `Time` задается время начала встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-636">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="18cff-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="18cff-637">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="18cff-638">Получает или задает описание, которое отображается в поле темы элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-638">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="18cff-639">Свойство `subject` получает или задает всю тему элемента для отправки с почтового сервера.</span><span class="sxs-lookup"><span data-stu-id="18cff-639">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-640">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-640">Read mode</span></span>

<span data-ttu-id="18cff-p131">Свойство `subject` возвращает строку. С помощью свойства [`normalizedSubject`](#normalizedsubject-string) можно получить тему без начальных префиксов, таких как `RE:` и `FW:`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="18cff-643">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-643">Compose mode</span></span>

<span data-ttu-id="18cff-644">Свойство `subject` возвращает объект `Subject`, который предоставляет методы для получения и задания темы.</span><span class="sxs-lookup"><span data-stu-id="18cff-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="18cff-645">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-645">Type:</span></span>

*   <span data-ttu-id="18cff-646">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="18cff-646">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-647">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-647">Requirements</span></span>

|<span data-ttu-id="18cff-648">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-648">Requirement</span></span>|<span data-ttu-id="18cff-649">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-650">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-651">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-651">1.0</span></span>|
|[<span data-ttu-id="18cff-652">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-653">ReadItem</span></span>|
|[<span data-ttu-id="18cff-654">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-655">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-655">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="18cff-656">Чтобы: массив. <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[получателей](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="18cff-657">Предоставляет доступ к получателей в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="18cff-658">Уровень доступа и тип объекта, зависит от режима текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="18cff-659">Режим чтения</span><span class="sxs-lookup"><span data-stu-id="18cff-659">Read mode</span></span>

<span data-ttu-id="18cff-p133">Свойство `to` возвращает массив, содержащий объект `EmailAddressDetails` для каждого получателя в строке **Кому** сообщения. Коллекция может включать не более 100 элементов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="18cff-662">Режим создания</span><span class="sxs-lookup"><span data-stu-id="18cff-662">Compose mode</span></span>

<span data-ttu-id="18cff-663">`to` Возвращает свойство `Recipients` объект, предоставляющий методы для получения или обновления получателей в строке **Кому** сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="18cff-664">Тип:</span><span class="sxs-lookup"><span data-stu-id="18cff-664">Type:</span></span>

*   <span data-ttu-id="18cff-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="18cff-665">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-666">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-666">Requirements</span></span>

|<span data-ttu-id="18cff-667">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-667">Requirement</span></span>|<span data-ttu-id="18cff-668">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-669">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-669">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-670">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-670">1.0</span></span>|
|[<span data-ttu-id="18cff-671">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-672">ReadItem</span></span>|
|[<span data-ttu-id="18cff-673">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-674">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-674">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-675">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-675">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="18cff-676">Методы</span><span class="sxs-lookup"><span data-stu-id="18cff-676">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="18cff-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18cff-677">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="18cff-678">Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="18cff-678">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="18cff-679">Метод `addFileAttachmentAsync` передает файл по указанному универсальному коду ресурса (URI) и вкладывает его в элемент в форме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-679">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="18cff-680">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="18cff-680">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-681">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-681">Parameters:</span></span>
|<span data-ttu-id="18cff-682">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-682">Name</span></span>|<span data-ttu-id="18cff-683">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-683">Type</span></span>|<span data-ttu-id="18cff-684">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-684">Attributes</span></span>|<span data-ttu-id="18cff-685">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-685">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="18cff-686">String</span><span class="sxs-lookup"><span data-stu-id="18cff-686">String</span></span>||<span data-ttu-id="18cff-p134">Универсальный код ресурса (URI), представляющий расположение файла, который нужно вложить в сообщение или встречу. Максимальная длина — 2048 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="18cff-689">String</span><span class="sxs-lookup"><span data-stu-id="18cff-689">String</span></span>||<span data-ttu-id="18cff-p135">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="18cff-692">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-692">Object</span></span>|<span data-ttu-id="18cff-693">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-693">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-694">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-694">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-695">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-695">Object</span></span>|<span data-ttu-id="18cff-696">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-696">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-697">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="18cff-697">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="18cff-698">Boolean</span><span class="sxs-lookup"><span data-stu-id="18cff-698">Boolean</span></span>|<span data-ttu-id="18cff-699">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-699">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-700">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-700">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="18cff-701">function</span><span class="sxs-lookup"><span data-stu-id="18cff-701">function</span></span>|<span data-ttu-id="18cff-702">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-702">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-703">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-703">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18cff-704">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18cff-704">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="18cff-705">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="18cff-705">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18cff-706">Ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-706">Errors</span></span>

|<span data-ttu-id="18cff-707">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-707">Error code</span></span>|<span data-ttu-id="18cff-708">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-708">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="18cff-709">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="18cff-709">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="18cff-710">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="18cff-710">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="18cff-711">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-711">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-712">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-712">Requirements</span></span>

|<span data-ttu-id="18cff-713">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-713">Requirement</span></span>|<span data-ttu-id="18cff-714">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-715">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-715">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-716">1.1</span><span class="sxs-lookup"><span data-stu-id="18cff-716">1.1</span></span>|
|[<span data-ttu-id="18cff-717">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-717">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-718">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-718">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-719">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-719">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-720">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-720">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="18cff-721">Примеры</span><span class="sxs-lookup"><span data-stu-id="18cff-721">Examples</span></span>

```js
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

<span data-ttu-id="18cff-722">В приведенном ниже примере файл изображения добавляется в качестве встроенного вложения, а ссылка на вложение добавляется в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-722">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="18cff-723">addFileAttachmentFromBase64Async (base64File, attachmentName, [параметры], [обратного вызова])</span><span class="sxs-lookup"><span data-stu-id="18cff-723">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="18cff-724">Добавляет файл из base64 кодирования в сообщение или встречу в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="18cff-724">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="18cff-725">`addFileAttachmentFromBase64Async` Метод загружает файл из кодировки base64 и подключает ее к элементу в форме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-725">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="18cff-726">Этот метод возвращает идентификатор вложения в объекте AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="18cff-726">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="18cff-727">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="18cff-727">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-728">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-728">Parameters:</span></span>
|<span data-ttu-id="18cff-729">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-729">Name</span></span>|<span data-ttu-id="18cff-730">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-730">Type</span></span>|<span data-ttu-id="18cff-731">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-731">Attributes</span></span>|<span data-ttu-id="18cff-732">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-732">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="18cff-733">String</span><span class="sxs-lookup"><span data-stu-id="18cff-733">String</span></span>||<span data-ttu-id="18cff-734">Контент, изображения или файла в электронной почте или событие добавляется в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="18cff-734">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="18cff-735">String</span><span class="sxs-lookup"><span data-stu-id="18cff-735">String</span></span>||<span data-ttu-id="18cff-p137">Имя вложения, которое отображается при передаче вложения. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="18cff-738">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-738">Object</span></span>|<span data-ttu-id="18cff-739">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-739">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-740">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-741">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-741">Object</span></span>|<span data-ttu-id="18cff-742">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-742">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-743">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="18cff-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="18cff-744">Boolean</span><span class="sxs-lookup"><span data-stu-id="18cff-744">Boolean</span></span>|<span data-ttu-id="18cff-745">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-745">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-746">Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="18cff-747">function</span><span class="sxs-lookup"><span data-stu-id="18cff-747">function</span></span>|<span data-ttu-id="18cff-748">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-748">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-749">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18cff-750">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18cff-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="18cff-751">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="18cff-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18cff-752">Ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-752">Errors</span></span>

|<span data-ttu-id="18cff-753">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-753">Error code</span></span>|<span data-ttu-id="18cff-754">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="18cff-755">Вложение превышает максимальный размер.</span><span class="sxs-lookup"><span data-stu-id="18cff-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="18cff-756">Расширение вложения не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="18cff-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="18cff-757">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-758">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-758">Requirements</span></span>

|<span data-ttu-id="18cff-759">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-759">Requirement</span></span>|<span data-ttu-id="18cff-760">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-761">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="18cff-761">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-762">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="18cff-762">Preview</span></span>|
|[<span data-ttu-id="18cff-763">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-765">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-766">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="18cff-767">Примеры</span><span class="sxs-lookup"><span data-stu-id="18cff-767">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="18cff-768">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18cff-768">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="18cff-769">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="18cff-769">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="18cff-770">В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="18cff-770">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-771">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-771">Parameters:</span></span>

| <span data-ttu-id="18cff-772">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-772">Name</span></span> | <span data-ttu-id="18cff-773">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-773">Type</span></span> | <span data-ttu-id="18cff-774">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-774">Attributes</span></span> | <span data-ttu-id="18cff-775">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-775">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="18cff-776">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="18cff-776">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="18cff-777">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="18cff-777">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="18cff-778">Function</span><span class="sxs-lookup"><span data-stu-id="18cff-778">Function</span></span> || <span data-ttu-id="18cff-p138">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="18cff-782">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-782">Object</span></span> | <span data-ttu-id="18cff-783">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-783">&lt;optional&gt;</span></span> | <span data-ttu-id="18cff-784">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-784">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="18cff-785">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-785">Object</span></span> | <span data-ttu-id="18cff-786">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-786">&lt;optional&gt;</span></span> | <span data-ttu-id="18cff-787">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-787">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="18cff-788">function</span><span class="sxs-lookup"><span data-stu-id="18cff-788">function</span></span>| <span data-ttu-id="18cff-789">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-789">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-790">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-791">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-791">Requirements</span></span>

|<span data-ttu-id="18cff-792">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-792">Requirement</span></span>| <span data-ttu-id="18cff-793">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-793">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-794">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-794">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18cff-795">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-795">1.7</span></span> |
|[<span data-ttu-id="18cff-796">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-796">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18cff-797">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-797">ReadItem</span></span> |
|[<span data-ttu-id="18cff-798">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-798">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18cff-799">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-799">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="18cff-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18cff-800">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="18cff-801">Добавляет к сообщению элемент Exchange, например сообщение, в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="18cff-801">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="18cff-p139">С помощью метода `addItemAttachmentAsync` можно в элемент формы создания вложить элемент с указанным идентификатором Exchange. Если указать метод обратного вызова, то этот метод вызывается с помощью параметра `asyncResult`, который содержит идентификатор вложения или код, указывающий на ошибки, которые произошли при вложении элемента. При необходимости можно использовать параметр `options` для передачи сведений о состоянии методу обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="18cff-805">Идентификатор можно последовательно использовать с методом [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback), чтобы удалить вложение, добавленное во время текущего сеанса.</span><span class="sxs-lookup"><span data-stu-id="18cff-805">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="18cff-806">Если надстройки Office работает в Outlook Web App, `addItemAttachmentAsync` метод могут прикреплять элементов для элементов, отличных от элемента, который вы изменяете; Однако это не поддерживается и не рекомендуется.</span><span class="sxs-lookup"><span data-stu-id="18cff-806">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-807">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-807">Parameters:</span></span>

|<span data-ttu-id="18cff-808">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-808">Name</span></span>|<span data-ttu-id="18cff-809">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-809">Type</span></span>|<span data-ttu-id="18cff-810">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-810">Attributes</span></span>|<span data-ttu-id="18cff-811">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-811">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="18cff-812">String</span><span class="sxs-lookup"><span data-stu-id="18cff-812">String</span></span>||<span data-ttu-id="18cff-p140">Идентификатор Exchange для вкладываемого элемента. Максимальная длина — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="18cff-815">String</span><span class="sxs-lookup"><span data-stu-id="18cff-815">String</span></span>||<span data-ttu-id="18cff-p141">Тема вкладываемого элемента. Максимальная длина — 255 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="18cff-818">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-818">Object</span></span>|<span data-ttu-id="18cff-819">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-819">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-820">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-820">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-821">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-821">Object</span></span>|<span data-ttu-id="18cff-822">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-822">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-823">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-823">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-824">function</span><span class="sxs-lookup"><span data-stu-id="18cff-824">function</span></span>|<span data-ttu-id="18cff-825">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-825">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-826">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-826">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18cff-827">После успешного выполнения идентификатор вложения будет представлен в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18cff-827">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="18cff-828">Если добавить вложение не удастся, объект `asyncResult` будет содержать объект `Error` с описанием ошибки.</span><span class="sxs-lookup"><span data-stu-id="18cff-828">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18cff-829">Ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-829">Errors</span></span>

|<span data-ttu-id="18cff-830">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-830">Error code</span></span>|<span data-ttu-id="18cff-831">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-831">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="18cff-832">Сообщение или встреча содержат слишком много вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-832">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-833">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-833">Requirements</span></span>

|<span data-ttu-id="18cff-834">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-834">Requirement</span></span>|<span data-ttu-id="18cff-835">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-836">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-837">1.1</span><span class="sxs-lookup"><span data-stu-id="18cff-837">1.1</span></span>|
|[<span data-ttu-id="18cff-838">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-839">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-839">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-840">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-841">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-841">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-842">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-842">Example</span></span>

<span data-ttu-id="18cff-843">В следующем примере существующий элемент Outlook добавляется в виде вложения с именем `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="18cff-843">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="18cff-844">close()</span><span class="sxs-lookup"><span data-stu-id="18cff-844">close()</span></span>

<span data-ttu-id="18cff-845">Закрывает текущий создаваемый элемент.</span><span class="sxs-lookup"><span data-stu-id="18cff-845">Closes the current item that is being composed.</span></span>

<span data-ttu-id="18cff-p142">Работа метода `close` зависит от текущего состояния создаваемого элемента. Если элемент содержит несохраненные изменения, клиент предложит пользователю сохранить или отклонить их либо отменить действие закрытия.</span><span class="sxs-lookup"><span data-stu-id="18cff-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-848">В Outlook в Интернете, если элемент является ли он встречей, и он ранее был сохранен с помощью `saveAsync`, то пользователю будет предложено сохранение, удаление или Отмена даже в том случае, если изменений внесено не было с элемента последнего сохранения.</span><span class="sxs-lookup"><span data-stu-id="18cff-848">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="18cff-849">Если в клиенте Outlook для настольных ПК сообщение представляет собой ответ в тексте, метод `close` не работает.</span><span class="sxs-lookup"><span data-stu-id="18cff-849">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-850">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-850">Requirements</span></span>

|<span data-ttu-id="18cff-851">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-851">Requirement</span></span>|<span data-ttu-id="18cff-852">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-853">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-853">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-854">1.3</span><span class="sxs-lookup"><span data-stu-id="18cff-854">1.3</span></span>|
|[<span data-ttu-id="18cff-855">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-855">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-856">Restricted</span><span class="sxs-lookup"><span data-stu-id="18cff-856">Restricted</span></span>|
|[<span data-ttu-id="18cff-857">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-857">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-858">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-858">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="18cff-859">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="18cff-859">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="18cff-860">Отображает форму ответа, включающую отправителя и всех получателей выбранного сообщения или организатора и всех участников выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-860">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-861">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-861">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-862">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="18cff-862">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="18cff-863">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyAllForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="18cff-863">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="18cff-p143">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="18cff-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-867">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-867">Parameters:</span></span>

|<span data-ttu-id="18cff-868">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-868">Name</span></span>|<span data-ttu-id="18cff-869">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-869">Type</span></span>|<span data-ttu-id="18cff-870">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-870">Attributes</span></span>|<span data-ttu-id="18cff-871">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-871">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="18cff-872">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="18cff-872">String &#124; Object</span></span>||<span data-ttu-id="18cff-p144">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="18cff-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="18cff-875">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="18cff-875">**OR**</span></span><br/><span data-ttu-id="18cff-p145">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="18cff-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="18cff-878">String</span><span class="sxs-lookup"><span data-stu-id="18cff-878">String</span></span>|<span data-ttu-id="18cff-879">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-879">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-p146">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="18cff-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="18cff-882">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-882">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="18cff-883">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-883">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-884">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="18cff-884">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="18cff-885">String</span><span class="sxs-lookup"><span data-stu-id="18cff-885">String</span></span>||<span data-ttu-id="18cff-p147">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="18cff-888">String</span><span class="sxs-lookup"><span data-stu-id="18cff-888">String</span></span>||<span data-ttu-id="18cff-889">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-889">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="18cff-890">String</span><span class="sxs-lookup"><span data-stu-id="18cff-890">String</span></span>||<span data-ttu-id="18cff-p148">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="18cff-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="18cff-893">Boolean</span><span class="sxs-lookup"><span data-stu-id="18cff-893">Boolean</span></span>||<span data-ttu-id="18cff-p149">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="18cff-896">String</span><span class="sxs-lookup"><span data-stu-id="18cff-896">String</span></span>||<span data-ttu-id="18cff-p150">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="18cff-900">function</span><span class="sxs-lookup"><span data-stu-id="18cff-900">function</span></span>|<span data-ttu-id="18cff-901">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-901">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-902">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-902">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-903">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-903">Requirements</span></span>

|<span data-ttu-id="18cff-904">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-904">Requirement</span></span>|<span data-ttu-id="18cff-905">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-906">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-906">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-907">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-907">1.0</span></span>|
|[<span data-ttu-id="18cff-908">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-909">ReadItem</span></span>|
|[<span data-ttu-id="18cff-910">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-911">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-911">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="18cff-912">Примеры</span><span class="sxs-lookup"><span data-stu-id="18cff-912">Examples</span></span>

<span data-ttu-id="18cff-913">Приведенный ниже код передает строку в функцию `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="18cff-913">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="18cff-914">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-914">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="18cff-915">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-915">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="18cff-916">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="18cff-916">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="18cff-917">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="18cff-917">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="18cff-918">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="18cff-918">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="18cff-919">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="18cff-919">displayReplyForm(formData)</span></span>

<span data-ttu-id="18cff-920">Отображает форму ответа, включающую только отправителя выбранного сообщения или организатора выбранной встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-920">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-921">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-921">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-922">В Outlook Web App форма ответа отображается в виде всплывающей формы в представлении с 3 либо 1 или 2 колонками.</span><span class="sxs-lookup"><span data-stu-id="18cff-922">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="18cff-923">Если любой строковый параметр превышает указанные для него ограничения, `displayReplyForm` возвращает исключение.</span><span class="sxs-lookup"><span data-stu-id="18cff-923">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="18cff-p151">Если в параметре `formData.attachments` указаны вложения, Outlook и Outlook Web App пытаются скачать их и вложить в форму ответа. Если какие-либо вложения добавить не удается, в форме отображается сообщение об ошибке. Если сообщения об ошибках не предусмотрены, то они не отображаются.</span><span class="sxs-lookup"><span data-stu-id="18cff-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-927">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-927">Parameters:</span></span>

|<span data-ttu-id="18cff-928">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-928">Name</span></span>|<span data-ttu-id="18cff-929">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-929">Type</span></span>|<span data-ttu-id="18cff-930">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-930">Attributes</span></span>|<span data-ttu-id="18cff-931">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-931">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="18cff-932">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="18cff-932">String &#124; Object</span></span>||<span data-ttu-id="18cff-p152">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="18cff-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="18cff-935">**ИЛИ**</span><span class="sxs-lookup"><span data-stu-id="18cff-935">**OR**</span></span><br/><span data-ttu-id="18cff-p153">Объект, который содержит текст или данные вложения и функцию обратного вызова. Ниже представлено определение этого объекта.</span><span class="sxs-lookup"><span data-stu-id="18cff-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="18cff-938">String</span><span class="sxs-lookup"><span data-stu-id="18cff-938">String</span></span>|<span data-ttu-id="18cff-939">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-939">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-p154">Строка, содержащая текст и HTML-код, представляющие собой основной текст формы ответа. Максимальный размер строки — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="18cff-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="18cff-942">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-942">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="18cff-943">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-943">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-944">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="18cff-944">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="18cff-945">String</span><span class="sxs-lookup"><span data-stu-id="18cff-945">String</span></span>||<span data-ttu-id="18cff-p155">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="18cff-948">String</span><span class="sxs-lookup"><span data-stu-id="18cff-948">String</span></span>||<span data-ttu-id="18cff-949">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-949">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="18cff-950">String</span><span class="sxs-lookup"><span data-stu-id="18cff-950">String</span></span>||<span data-ttu-id="18cff-p156">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="18cff-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="18cff-953">Boolean</span><span class="sxs-lookup"><span data-stu-id="18cff-953">Boolean</span></span>||<span data-ttu-id="18cff-p157">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="18cff-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="18cff-956">String</span><span class="sxs-lookup"><span data-stu-id="18cff-956">String</span></span>||<span data-ttu-id="18cff-p158">Используется, только если свойству `type` задано значение `item`. Идентификатор вложения EWS. Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="18cff-960">function</span><span class="sxs-lookup"><span data-stu-id="18cff-960">function</span></span>|<span data-ttu-id="18cff-961">&lt;Необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-961">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-962">По завершении работы метода функция, переданная параметру `callback`, вызывается с помощью одного параметра `asyncResult`, представляющего собой объект [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-962">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-963">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-963">Requirements</span></span>

|<span data-ttu-id="18cff-964">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-964">Requirement</span></span>|<span data-ttu-id="18cff-965">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-966">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-966">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-967">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-967">1.0</span></span>|
|[<span data-ttu-id="18cff-968">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-968">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-969">ReadItem</span></span>|
|[<span data-ttu-id="18cff-970">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-970">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-971">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-971">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="18cff-972">Примеры</span><span class="sxs-lookup"><span data-stu-id="18cff-972">Examples</span></span>

<span data-ttu-id="18cff-973">Приведенный ниже код передает строку в функцию `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="18cff-973">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="18cff-974">Ответ с пустым текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-974">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="18cff-975">Ответ только с текстом сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-975">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="18cff-976">Ответ с текстом сообщения и вложенным файлом.</span><span class="sxs-lookup"><span data-stu-id="18cff-976">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="18cff-977">Ответ с текстом сообщения и вложенным элементом.</span><span class="sxs-lookup"><span data-stu-id="18cff-977">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="18cff-978">Ответ с текстом сообщения, вложенным файлом, вложенным элементом и обратным вызовом.</span><span class="sxs-lookup"><span data-stu-id="18cff-978">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="18cff-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="18cff-979">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="18cff-980">Возвращает сущности, обнаруженные в тело выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-980">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-981">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-981">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-982">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-982">Requirements</span></span>

|<span data-ttu-id="18cff-983">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-983">Requirement</span></span>|<span data-ttu-id="18cff-984">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-985">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-985">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-986">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-986">1.0</span></span>|
|[<span data-ttu-id="18cff-987">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-988">ReadItem</span></span>|
|[<span data-ttu-id="18cff-989">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-990">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-990">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-991">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-991">Returns:</span></span>

<span data-ttu-id="18cff-992">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="18cff-992">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="18cff-993">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-993">Example</span></span>

<span data-ttu-id="18cff-994">Этот пример ссылается сущностей контакты в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-994">The following example accesses the contacts entities in the current item's body.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="18cff-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="18cff-995">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="18cff-996">Получает массив всех сущностей указанного типа, обнаруженных в тело выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-996">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-997">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-997">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-998">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-998">Parameters:</span></span>

|<span data-ttu-id="18cff-999">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-999">Name</span></span>|<span data-ttu-id="18cff-1000">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1000">Type</span></span>|<span data-ttu-id="18cff-1001">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1001">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="18cff-1002">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="18cff-1002">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="18cff-1003">Одно из значений перечисления EntityType.</span><span class="sxs-lookup"><span data-stu-id="18cff-1003">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1004">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1004">Requirements</span></span>

|<span data-ttu-id="18cff-1005">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1005">Requirement</span></span>|<span data-ttu-id="18cff-1006">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1007">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1007">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-1008">1.0</span></span>|
|[<span data-ttu-id="18cff-1009">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1010">Restricted</span><span class="sxs-lookup"><span data-stu-id="18cff-1010">Restricted</span></span>|
|[<span data-ttu-id="18cff-1011">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1012">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1012">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1013">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1013">Returns:</span></span>

<span data-ttu-id="18cff-1014">Если значение, переданное в `entityType`, не является допустимым членом перечисления `EntityType`, метод возвращает значение NULL.</span><span class="sxs-lookup"><span data-stu-id="18cff-1014">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="18cff-1015">Если сущности указанного типа отсутствуют в основной текст элемента, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="18cff-1015">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="18cff-1016">В противном случае тип объектов в возвращаемом массиве зависит от типа сущности, запрошенной в параметре `entityType`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1016">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="18cff-1017">Хотя минимальный уровень разрешений для использования этого метода — **Restricted**, для некоторых типов сущностей требуется доступ на уровне **ReadItem**, как указано в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="18cff-1017">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="18cff-1018">Значение параметра `entityType`</span><span class="sxs-lookup"><span data-stu-id="18cff-1018">Value of `entityType`</span></span>|<span data-ttu-id="18cff-1019">Тип объектов в возвращаемом массиве</span><span class="sxs-lookup"><span data-stu-id="18cff-1019">Type of objects in returned array</span></span>|<span data-ttu-id="18cff-1020">Необходимый уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1020">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="18cff-1021">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1021">String</span></span>|<span data-ttu-id="18cff-1022">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18cff-1022">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="18cff-1023">Contact</span><span class="sxs-lookup"><span data-stu-id="18cff-1023">Contact</span></span>|<span data-ttu-id="18cff-1024">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18cff-1024">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="18cff-1025">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1025">String</span></span>|<span data-ttu-id="18cff-1026">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18cff-1026">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="18cff-1027">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="18cff-1027">MeetingSuggestion</span></span>|<span data-ttu-id="18cff-1028">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18cff-1028">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="18cff-1029">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="18cff-1029">PhoneNumber</span></span>|<span data-ttu-id="18cff-1030">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18cff-1030">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="18cff-1031">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="18cff-1031">TaskSuggestion</span></span>|<span data-ttu-id="18cff-1032">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="18cff-1032">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="18cff-1033">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1033">String</span></span>|<span data-ttu-id="18cff-1034">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="18cff-1034">**Restricted**</span></span>|

<span data-ttu-id="18cff-1035">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="18cff-1035">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="18cff-1036">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1036">Example</span></span>

<span data-ttu-id="18cff-1037">Следующем примере показано, как получить доступ к массив строк, представляющих почтовых адресов в тексте текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-1037">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="18cff-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="18cff-1038">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="18cff-1039">Возвращает известные сущности в выбранном элементе, которые проходят через именованный фильтр, определяемый в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="18cff-1039">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1040">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-1040">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-1041">Метод `getFilteredEntitiesByName` возвращает сущности, соответствующие регулярному выражению, которое определяется в элементе правила [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) в XML-файле манифеста, с использованием указанного значения элемента `FilterName`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1041">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1042">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1042">Parameters:</span></span>

|<span data-ttu-id="18cff-1043">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1043">Name</span></span>|<span data-ttu-id="18cff-1044">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1044">Type</span></span>|<span data-ttu-id="18cff-1045">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1045">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="18cff-1046">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1046">String</span></span>|<span data-ttu-id="18cff-1047">Имя элемента правила `ItemHasKnownEntity`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="18cff-1047">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1048">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1048">Requirements</span></span>

|<span data-ttu-id="18cff-1049">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1049">Requirement</span></span>|<span data-ttu-id="18cff-1050">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1051">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1051">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-1052">1.0</span></span>|
|[<span data-ttu-id="18cff-1053">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1054">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1055">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1056">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1056">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1057">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1057">Returns:</span></span>

<span data-ttu-id="18cff-p160">Если в манифесте нет элемента `ItemHasKnownEntity` со значением `FilterName`, соответствующим параметру `name`, метод возвращает `null`. Если параметр `name` соответствует элементу `ItemHasKnownEntity` в манифесте, но при этом в текущем элементе нет соответствующих сущностей, метод возвращает пустой массив.</span><span class="sxs-lookup"><span data-stu-id="18cff-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="18cff-1060">Тип: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="18cff-1060">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="18cff-1061">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18cff-1061">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="18cff-1062">Получает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="18cff-1062">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1063">Этот метод поддерживается только по Outlook 2016 или более поздней версии для Windows (более поздней, чем 16.0.8413.1000 версии Click-to-Run) и Outlook в Интернете для Office 365.</span><span class="sxs-lookup"><span data-stu-id="18cff-1063">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1064">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1064">Parameters:</span></span>
|<span data-ttu-id="18cff-1065">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1065">Name</span></span>|<span data-ttu-id="18cff-1066">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1066">Type</span></span>|<span data-ttu-id="18cff-1067">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1067">Attributes</span></span>|<span data-ttu-id="18cff-1068">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1068">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="18cff-1069">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-1069">Object</span></span>|<span data-ttu-id="18cff-1070">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1070">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1071">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1071">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1072">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1072">Object</span></span>|<span data-ttu-id="18cff-1073">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1073">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1074">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1074">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-1075">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1075">function</span></span>|<span data-ttu-id="18cff-1076">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1076">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1077">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1077">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18cff-1078">В случае успешного выполнения инициализации данных предоставляются в `asyncResult.value` свойства в виде строки.</span><span class="sxs-lookup"><span data-stu-id="18cff-1078">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="18cff-1079">Если контекст инициализации отсутствует, объект `asyncResult` будет содержать объект `Error`, одному свойству которого (`code`) будет присвоено значение `9020`, а другому (`name`) — значение `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1079">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1080">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1080">Requirements</span></span>

|<span data-ttu-id="18cff-1081">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1081">Requirement</span></span>|<span data-ttu-id="18cff-1082">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1083">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="18cff-1083">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1084">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="18cff-1084">Preview</span></span>|
|[<span data-ttu-id="18cff-1085">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1085">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1086">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1087">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1087">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1088">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1088">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-1089">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1089">Example</span></span>

```
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="18cff-1090">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="18cff-1090">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="18cff-1091">Возвращает строковые значения в выбранном элементе, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="18cff-1091">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1092">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-p161">Метод `getRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="18cff-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="18cff-1096">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1096">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="18cff-1097">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1097">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="18cff-p162">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="18cff-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-1101">Requirements</span><span class="sxs-lookup"><span data-stu-id="18cff-1101">Requirements</span></span>

|<span data-ttu-id="18cff-1102">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1102">Requirement</span></span>|<span data-ttu-id="18cff-1103">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1103">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1104">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1104">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1105">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-1105">1.0</span></span>|
|[<span data-ttu-id="18cff-1106">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1106">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1107">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1107">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1108">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1108">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1109">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1109">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1110">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1110">Returns:</span></span>

<span data-ttu-id="18cff-p163">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="18cff-1113">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="18cff-1113">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18cff-1114">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1114">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18cff-1115">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1115">Example</span></span>

<span data-ttu-id="18cff-1116">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="18cff-1116">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="18cff-1117">getRegExMatchesByName(name) пункты (допускает значение NULL) {массива. < String >}</span><span class="sxs-lookup"><span data-stu-id="18cff-1117">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="18cff-1118">Возвращает строковые значения в выбранном элементе, которые соответствуют именованному регулярному выражению, определенному в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="18cff-1118">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1119">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-1119">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-1120">Метод `getRegExMatchesByName` возвращает строки, соответствующие регулярному выражению, которое определяется в элементе правила `ItemHasRegularExpressionMatch` в XML-файле манифеста, с использованием указанного значения элемента `RegExName`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1120">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="18cff-p164">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты.</span><span class="sxs-lookup"><span data-stu-id="18cff-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1123">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1123">Parameters:</span></span>

|<span data-ttu-id="18cff-1124">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1124">Name</span></span>|<span data-ttu-id="18cff-1125">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1125">Type</span></span>|<span data-ttu-id="18cff-1126">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1126">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="18cff-1127">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1127">String</span></span>|<span data-ttu-id="18cff-1128">Имя элемента правила `ItemHasRegularExpressionMatch`, определяющее соответствующий фильтр.</span><span class="sxs-lookup"><span data-stu-id="18cff-1128">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1129">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1129">Requirements</span></span>

|<span data-ttu-id="18cff-1130">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1130">Requirement</span></span>|<span data-ttu-id="18cff-1131">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1133">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-1133">1.0</span></span>|
|[<span data-ttu-id="18cff-1134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1135">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1137">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1137">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1138">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1138">Returns:</span></span>

<span data-ttu-id="18cff-1139">Массив строк, соответствующих регулярному выражению, определяемому в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="18cff-1139">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="18cff-1140">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="18cff-1140">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18cff-1141">Массив. < String ></span><span class="sxs-lookup"><span data-stu-id="18cff-1141">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18cff-1142">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1142">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="18cff-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="18cff-1143">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="18cff-1144">Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-1144">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="18cff-p165">Если выделенный фрагмент отсутствует, но курсор находится в тексте или теме, метод возвращает значение NULL для выбранных данных. Если выбраны не текст и не тема, метод возвращает ошибку `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1147">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1147">Parameters:</span></span>

|<span data-ttu-id="18cff-1148">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1148">Name</span></span>|<span data-ttu-id="18cff-1149">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1149">Type</span></span>|<span data-ttu-id="18cff-1150">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1150">Attributes</span></span>|<span data-ttu-id="18cff-1151">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1151">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="18cff-1152">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="18cff-1152">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="18cff-p166">Запрашивает формат данных. Если задано значение Text, метод возвращает обычный текст как строку, удаляя все имеющиеся HTML-теги. Если задано значение HTML, метод возвращает выделенный текст (обычный текст или HTML).</span><span class="sxs-lookup"><span data-stu-id="18cff-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="18cff-1156">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1156">Object</span></span>|<span data-ttu-id="18cff-1157">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1158">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1158">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1159">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1159">Object</span></span>|<span data-ttu-id="18cff-1160">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1161">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1161">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-1162">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1162">function</span></span>||<span data-ttu-id="18cff-1163">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1163">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18cff-1164">Чтобы получить доступ к выбранным данным из метода обратного вызова, вызовите `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1164">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="18cff-1165">Для доступа к свойству источника, выделение, поступающих из источников, вызовите `asyncResult.value.sourceProperty`, который может быть либо `body` или `subject`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1165">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1166">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1166">Requirements</span></span>

|<span data-ttu-id="18cff-1167">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1167">Requirement</span></span>|<span data-ttu-id="18cff-1168">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1168">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1169">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="18cff-1169">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1170">1.2</span><span class="sxs-lookup"><span data-stu-id="18cff-1170">1.2</span></span>|
|[<span data-ttu-id="18cff-1171">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1171">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1172">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1172">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-1173">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1173">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1174">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-1174">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1175">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1175">Returns:</span></span>

<span data-ttu-id="18cff-1176">Выбранные данные в виде строки с форматом, определенным в параметре `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1176">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="18cff-1177">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="18cff-1177">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="18cff-1178">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1178">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="18cff-1179">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1179">Example</span></span>

```
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="18cff-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="18cff-1180">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="18cff-p168">Возвращает сущности, найденные в выделенном совпадении, выбранном пользователем. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="18cff-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1183">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-1183">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-1184">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1184">Requirements</span></span>

|<span data-ttu-id="18cff-1185">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1185">Requirement</span></span>|<span data-ttu-id="18cff-1186">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1186">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1187">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1188">1.6</span><span class="sxs-lookup"><span data-stu-id="18cff-1188">1.6</span></span>|
|[<span data-ttu-id="18cff-1189">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1189">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1190">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1192">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1192">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1193">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1193">Returns:</span></span>

<span data-ttu-id="18cff-1194">Тип: [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="18cff-1194">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="18cff-1195">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1195">Example</span></span>

<span data-ttu-id="18cff-1196">В приведенном ниже примере показано, как получить доступ к сущностям адресов в выделенном совпадении, выбранном пользователем.</span><span class="sxs-lookup"><span data-stu-id="18cff-1196">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="18cff-1197">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="18cff-1197">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="18cff-p169">Возвращает строковые значения в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к [контекстным надстройкам](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="18cff-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1200">Этот метод не поддерживается в Outlook для операций ввода-вывода или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="18cff-1200">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="18cff-p170">Метод `getSelectedRegExMatches` возвращает строки, соответствующие регулярному выражению, которое определяется в каждом элементе правила `ItemHasRegularExpressionMatch` или `ItemHasKnownEntity` в XML-файле манифеста. Для правила `ItemHasRegularExpressionMatch` соответствующую строку должно содержать свойство элемента, указанного этим правилом. Простой тип `PropertyName` определяет поддерживаемые свойства.</span><span class="sxs-lookup"><span data-stu-id="18cff-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="18cff-1204">Например, рассмотрим манифест надстройки, который содержит указанный ниже элемент `Rule`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1204">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="18cff-1205">Объект, возвращаемый методом `getRegExMatches`, будет содержать два свойства: `fruits` и `veggies`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1205">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="18cff-p171">Если вы указываете правило `ItemHasRegularExpressionMatch` для свойства текста элемента, регулярное выражение должно дальше фильтровать текст, а не пытаться вернуть весь текст элемента. Использование регулярного выражения, такого как `.*`, для получения всего текста элемента не всегда приносит ожидаемые результаты. Лучше используйте метод [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) для этого.</span><span class="sxs-lookup"><span data-stu-id="18cff-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="18cff-1209">Requirements</span><span class="sxs-lookup"><span data-stu-id="18cff-1209">Requirements</span></span>

|<span data-ttu-id="18cff-1210">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1210">Requirement</span></span>|<span data-ttu-id="18cff-1211">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1212">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1212">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1213">1.6</span><span class="sxs-lookup"><span data-stu-id="18cff-1213">1.6</span></span>|
|[<span data-ttu-id="18cff-1214">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1214">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1215">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1217">Чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1217">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="18cff-1218">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="18cff-1218">Returns:</span></span>

<span data-ttu-id="18cff-p172">Объект, содержащий массив строк, которые соответствуют регулярным выражениям, определяемым в XML-файле манифеста. Имя каждого массива равно соответствующему значению атрибута `RegExName` подходящего правила `ItemHasRegularExpressionMatch` или атрибута `FilterName` соответствующего правила `ItemHasKnownEntity`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="18cff-1221">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1221">Example</span></span>

<span data-ttu-id="18cff-1222">В приведенном ниже примере показано, как получить доступ к массиву совпадений с элементами `fruits` и `veggies` правил активации регулярных выражений, указанными в манифесте.</span><span class="sxs-lookup"><span data-stu-id="18cff-1222">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="18cff-1223">getSharedPropertiesAsync ([параметры] обратного вызова)</span><span class="sxs-lookup"><span data-stu-id="18cff-1223">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="18cff-1224">Получает свойства выбранной встречи или сообщения в общей папке, календаря или почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="18cff-1224">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1225">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1225">Parameters:</span></span>

|<span data-ttu-id="18cff-1226">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1226">Name</span></span>|<span data-ttu-id="18cff-1227">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1227">Type</span></span>|<span data-ttu-id="18cff-1228">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1228">Attributes</span></span>|<span data-ttu-id="18cff-1229">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1229">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="18cff-1230">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-1230">Object</span></span>|<span data-ttu-id="18cff-1231">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1232">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1232">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1233">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1233">Object</span></span>|<span data-ttu-id="18cff-1234">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1234">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1235">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1235">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-1236">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1236">function</span></span>||<span data-ttu-id="18cff-1237">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1237">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18cff-1238">Общие свойства предоставляются как [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) объект в `asyncResult.value` свойство.</span><span class="sxs-lookup"><span data-stu-id="18cff-1238">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="18cff-1239">Этот объект можно использовать для получения общего свойства элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-1239">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1240">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1240">Requirements</span></span>

|<span data-ttu-id="18cff-1241">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1241">Requirement</span></span>|<span data-ttu-id="18cff-1242">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1243">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="18cff-1243">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1244">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="18cff-1244">Preview</span></span>|
|[<span data-ttu-id="18cff-1245">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1246">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1247">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1248">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1248">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-1249">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1249">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="18cff-1250">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="18cff-1250">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="18cff-1251">Асинхронно загружает настраиваемые свойства для надстройки для выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-1251">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="18cff-p174">Настраиваемые свойства сохраняются в виде пар "ключ-значение" для каждого приложения и каждого элемента. Этот метод возвращает объект `CustomProperties` при обратном вызове, который предоставляет методы для доступа к настраиваемым свойствам, характерным для текущего элемента и текущей надстройки. Настраиваемые свойства не шифруются для элемента, поэтому этот способ хранения не является безопасным.</span><span class="sxs-lookup"><span data-stu-id="18cff-p174">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1255">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1255">Parameters:</span></span>

|<span data-ttu-id="18cff-1256">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1256">Name</span></span>|<span data-ttu-id="18cff-1257">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1257">Type</span></span>|<span data-ttu-id="18cff-1258">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1258">Attributes</span></span>|<span data-ttu-id="18cff-1259">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1259">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="18cff-1260">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1260">function</span></span>||<span data-ttu-id="18cff-1261">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18cff-1262">Настраиваемые свойства предоставляются в виде объекта [`CustomProperties`](/javascript/api/outlook/office.customproperties) в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="18cff-1262">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="18cff-1263">Этот объект можно использовать для получения, задания и удаление настраиваемых свойств из элемента и сохранение изменений для настраиваемого свойства, задайте обратно на сервер.</span><span class="sxs-lookup"><span data-stu-id="18cff-1263">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="18cff-1264">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-1264">Object</span></span>|<span data-ttu-id="18cff-1265">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1266">Разработчики могут предоставлять любого объекта, которые следует получить доступ к в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1266">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="18cff-1267">Этот объект можно получить доступ с `asyncResult.asyncContext` в функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1267">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1268">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1268">Requirements</span></span>

|<span data-ttu-id="18cff-1269">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1269">Requirement</span></span>|<span data-ttu-id="18cff-1270">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1271">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1271">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="18cff-1272">1.0</span></span>|
|[<span data-ttu-id="18cff-1273">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1274">ReadItem</span></span>|
|[<span data-ttu-id="18cff-1275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1276">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-1277">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1277">Example</span></span>

<span data-ttu-id="18cff-p177">Приведенный ниже пример кода показывает, как асинхронно загружать настраиваемые свойства, характерные для текущего элемента, с помощью метода `loadCustomPropertiesAsync`. Этот пример также показывает, как сохранять эти свойства на сервере с помощью метода `CustomProperties.saveAsync`. После загрузки настраиваемых свойств в этом примере кода метод `CustomProperties.get` используется для считывания настраиваемого свойства `myProp`, метод `CustomProperties.set` — для записи настраиваемого свойства `otherProp`, а метод `saveAsync` — для сохранения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-p177">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="18cff-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="18cff-1281">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="18cff-1282">Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="18cff-1282">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="18cff-p178">Метод `removeAttachmentAsync` удаляет из элемента вложение с указанным идентификатором. Идентификатор вложения рекомендуется использовать для удаления вложения, только если оно добавлено тем же почтовым приложением в ходе текущего сеанса. В Outlook Web App и Outlook Web App для устройств идентификатор вложения действителен только в рамках одного сеанса. Сеанс завершается, когда пользователь закрывает приложение или начинает создавать элемент во встроенной форме, а затем переходит из формы в отдельное окно.</span><span class="sxs-lookup"><span data-stu-id="18cff-p178">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1287">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1287">Parameters:</span></span>

|<span data-ttu-id="18cff-1288">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1288">Name</span></span>|<span data-ttu-id="18cff-1289">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1289">Type</span></span>|<span data-ttu-id="18cff-1290">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1290">Attributes</span></span>|<span data-ttu-id="18cff-1291">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1291">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="18cff-1292">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1292">String</span></span>||<span data-ttu-id="18cff-p179">Идентификатор удаляемого вложения. Максимальная длина строки — 100 символов.</span><span class="sxs-lookup"><span data-stu-id="18cff-p179">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="18cff-1295">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1295">Object</span></span>|<span data-ttu-id="18cff-1296">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1297">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1297">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1298">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1298">Object</span></span>|<span data-ttu-id="18cff-1299">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1299">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1300">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1300">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-1301">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1301">function</span></span>|<span data-ttu-id="18cff-1302">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1303">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1303">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="18cff-1304">Если удалить вложение не удается, свойство `asyncResult.error` содержит код ошибки с указанием ее причины.</span><span class="sxs-lookup"><span data-stu-id="18cff-1304">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="18cff-1305">Ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-1305">Errors</span></span>

|<span data-ttu-id="18cff-1306">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="18cff-1306">Error code</span></span>|<span data-ttu-id="18cff-1307">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1307">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="18cff-1308">Идентификатор вложения не существует.</span><span class="sxs-lookup"><span data-stu-id="18cff-1308">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1309">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1309">Requirements</span></span>

|<span data-ttu-id="18cff-1310">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1310">Requirement</span></span>|<span data-ttu-id="18cff-1311">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1311">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1312">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1312">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1313">1.1</span><span class="sxs-lookup"><span data-stu-id="18cff-1313">1.1</span></span>|
|[<span data-ttu-id="18cff-1314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1315">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1315">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-1316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1317">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-1317">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-1318">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1318">Example</span></span>

<span data-ttu-id="18cff-1319">Указанный ниже код удаляет вложение с идентификатором "0".</span><span class="sxs-lookup"><span data-stu-id="18cff-1319">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="18cff-1320">removeHandlerAsync (тип события, обработчик, [параметры], [обратного вызова])</span><span class="sxs-lookup"><span data-stu-id="18cff-1320">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="18cff-1321">Удаляет обработчик событий для события, поддерживаемые.</span><span class="sxs-lookup"><span data-stu-id="18cff-1321">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="18cff-1322">В настоящее время поддерживаемые типы событий, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, и`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="18cff-1322">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1323">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1323">Parameters:</span></span>

| <span data-ttu-id="18cff-1324">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1324">Name</span></span> | <span data-ttu-id="18cff-1325">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1325">Type</span></span> | <span data-ttu-id="18cff-1326">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1326">Attributes</span></span> | <span data-ttu-id="18cff-1327">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1327">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="18cff-1328">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="18cff-1328">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="18cff-1329">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="18cff-1329">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="18cff-1330">Function</span><span class="sxs-lookup"><span data-stu-id="18cff-1330">Function</span></span> || <span data-ttu-id="18cff-p180">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p180">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="18cff-1334">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-1334">Object</span></span> | <span data-ttu-id="18cff-1335">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1335">&lt;optional&gt;</span></span> | <span data-ttu-id="18cff-1336">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1336">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="18cff-1337">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1337">Object</span></span> | <span data-ttu-id="18cff-1338">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1338">&lt;optional&gt;</span></span> | <span data-ttu-id="18cff-1339">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1339">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="18cff-1340">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1340">function</span></span>| <span data-ttu-id="18cff-1341">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1341">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1342">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1343">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1343">Requirements</span></span>

|<span data-ttu-id="18cff-1344">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1344">Requirement</span></span>| <span data-ttu-id="18cff-1345">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1345">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1346">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1346">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="18cff-1347">1.7</span><span class="sxs-lookup"><span data-stu-id="18cff-1347">1.7</span></span> |
|[<span data-ttu-id="18cff-1348">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1348">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="18cff-1349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1349">ReadItem</span></span> |
|[<span data-ttu-id="18cff-1350">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1350">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="18cff-1351">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="18cff-1351">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="18cff-1352">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="18cff-1352">saveAsync([options], callback)</span></span>

<span data-ttu-id="18cff-1353">Асинхронно сохраняет элемент.</span><span class="sxs-lookup"><span data-stu-id="18cff-1353">Asynchronously saves an item.</span></span>

<span data-ttu-id="18cff-p181">При вызове этот метод сохраняет текущее сообщение в виде черновика и возвращает идентификатор элемента с помощью метода обратного вызова. В Outlook Web App или интерактивном режиме Outlook этот элемент сохраняется на сервере. В Outlook в режиме кэширования этот элемент сохраняется в локальном кэше.</span><span class="sxs-lookup"><span data-stu-id="18cff-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1357">Если надстройка вызывает `saveAsync` элемент в режиме создания для получения `itemId` для использования с помощью веб-служб Exchange или интерфейса API REST, необходимо учитывать, что когда Outlook находится в режиме кэширования, он может занять некоторое время до элемента фактически синхронизируется с сервера.</span><span class="sxs-lookup"><span data-stu-id="18cff-1357">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="18cff-1358">Пока элемент синхронизирован с помощью `itemId` возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="18cff-1358">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="18cff-p183">Если метод `saveAsync` вызывается для встречи в режиме создания, она сохраняется как обычная встреча в календаре пользователя, а не как черновик. При сохранении новой встречи приглашения не отправляются. При сохранении существующей встречи уведомления отправляются добавленным или удаленным участникам.</span><span class="sxs-lookup"><span data-stu-id="18cff-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="18cff-1362">Следующие клиенты имеют по-разному для `saveAsync` для встреч в режиме создания:</span><span class="sxs-lookup"><span data-stu-id="18cff-1362">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="18cff-1363">Mac Outlook не поддерживает `saveAsync` на собрании в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-1363">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="18cff-1364">Вызов `saveAsync` собрания в Mac Outlook возвращает ошибку.</span><span class="sxs-lookup"><span data-stu-id="18cff-1364">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="18cff-1365">Outlook в Интернете всегда отправляет приглашение или обновления при `saveAsync` вызван на встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="18cff-1365">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1366">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1366">Parameters:</span></span>

|<span data-ttu-id="18cff-1367">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1367">Name</span></span>|<span data-ttu-id="18cff-1368">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1368">Type</span></span>|<span data-ttu-id="18cff-1369">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1369">Attributes</span></span>|<span data-ttu-id="18cff-1370">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1370">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="18cff-1371">Объект</span><span class="sxs-lookup"><span data-stu-id="18cff-1371">Object</span></span>|<span data-ttu-id="18cff-1372">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1372">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1373">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1373">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1374">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1374">Object</span></span>|<span data-ttu-id="18cff-1375">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1375">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1376">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="18cff-1376">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="18cff-1377">function</span><span class="sxs-lookup"><span data-stu-id="18cff-1377">function</span></span>||<span data-ttu-id="18cff-1378">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1378">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="18cff-1379">В случае успешного выполнения, идентификатор элемента представлен в `asyncResult.value` свойство.</span><span class="sxs-lookup"><span data-stu-id="18cff-1379">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1380">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1380">Requirements</span></span>

|<span data-ttu-id="18cff-1381">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1381">Requirement</span></span>|<span data-ttu-id="18cff-1382">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1382">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1383">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="18cff-1383">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1384">1.3</span><span class="sxs-lookup"><span data-stu-id="18cff-1384">1.3</span></span>|
|[<span data-ttu-id="18cff-1385">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1386">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1386">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-1387">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1388">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-1388">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="18cff-1389">Примеры</span><span class="sxs-lookup"><span data-stu-id="18cff-1389">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="18cff-p185">Ниже приведен пример параметра `result`, переданного функции обратного вызова. Свойство `value` содержит идентификатор элемента.</span><span class="sxs-lookup"><span data-stu-id="18cff-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="18cff-1392">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="18cff-1392">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="18cff-1393">Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="18cff-1393">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="18cff-p186">Метод `setSelectedDataAsync` вставляет указанную строку в местоположение курсора в теме или тексте элемента либо, если текст выделен в редакторе, он заменяет выделенный текст. Если курсор находится вне текста или темы элемента, возвращается ошибка. После вставки курсор помещается в конец вставленного содержимого.</span><span class="sxs-lookup"><span data-stu-id="18cff-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="18cff-1397">Параметры</span><span class="sxs-lookup"><span data-stu-id="18cff-1397">Parameters:</span></span>

|<span data-ttu-id="18cff-1398">Имя</span><span class="sxs-lookup"><span data-stu-id="18cff-1398">Name</span></span>|<span data-ttu-id="18cff-1399">Тип</span><span class="sxs-lookup"><span data-stu-id="18cff-1399">Type</span></span>|<span data-ttu-id="18cff-1400">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="18cff-1400">Attributes</span></span>|<span data-ttu-id="18cff-1401">Описание</span><span class="sxs-lookup"><span data-stu-id="18cff-1401">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="18cff-1402">String</span><span class="sxs-lookup"><span data-stu-id="18cff-1402">String</span></span>||<span data-ttu-id="18cff-p187">Вставляемые данные. Объем данных не должен превышать 1 000 000 символов. Если передано больше 1 000 000 символов, возвращается исключение `ArgumentOutOfRange`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="18cff-1406">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1406">Object</span></span>|<span data-ttu-id="18cff-1407">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1407">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1408">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="18cff-1408">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="18cff-1409">Object</span><span class="sxs-lookup"><span data-stu-id="18cff-1409">Object</span></span>|<span data-ttu-id="18cff-1410">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1410">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-1411">В методе обратного вызова разработчики могут указать любой объект, к которому необходимо получить доступ.</span><span class="sxs-lookup"><span data-stu-id="18cff-1411">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="18cff-1412">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="18cff-1412">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="18cff-1413">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="18cff-1413">&lt;optional&gt;</span></span>|<span data-ttu-id="18cff-p188">Если задано значение `text`, текущий стиль применяется в Outlook Web App и Outlook. Если поле представляет собой редактор HTML, вставляются только текстовые данные, даже если они имеют формат HTML.</span><span class="sxs-lookup"><span data-stu-id="18cff-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="18cff-p189">Если задано значение `html` и поле (не тема) поддерживает HTML, в Outlook Web App применяется текущий стиль, а в Outlook — стиль по умолчанию. Если поле является текстовым, возвращается ошибка `InvalidDataFormat`.</span><span class="sxs-lookup"><span data-stu-id="18cff-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="18cff-1418">Если свойство `coercionType` не задано, результат зависит от поля: если поле имеет формат HTML, используется текст в формате HTML, а если поле текстовое, применяется обычный текст.</span><span class="sxs-lookup"><span data-stu-id="18cff-1418">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="18cff-1419">функция</span><span class="sxs-lookup"><span data-stu-id="18cff-1419">function</span></span>||<span data-ttu-id="18cff-1420">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="18cff-1420">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="18cff-1421">Требования</span><span class="sxs-lookup"><span data-stu-id="18cff-1421">Requirements</span></span>

|<span data-ttu-id="18cff-1422">Requirement</span><span class="sxs-lookup"><span data-stu-id="18cff-1422">Requirement</span></span>|<span data-ttu-id="18cff-1423">Значение</span><span class="sxs-lookup"><span data-stu-id="18cff-1423">Value</span></span>|
|---|---|
|[<span data-ttu-id="18cff-1424">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="18cff-1424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="18cff-1425">1.2</span><span class="sxs-lookup"><span data-stu-id="18cff-1425">1.2</span></span>|
|[<span data-ttu-id="18cff-1426">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="18cff-1426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="18cff-1427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="18cff-1427">ReadWriteItem</span></span>|
|[<span data-ttu-id="18cff-1428">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="18cff-1428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="18cff-1429">Создание</span><span class="sxs-lookup"><span data-stu-id="18cff-1429">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="18cff-1430">Пример</span><span class="sxs-lookup"><span data-stu-id="18cff-1430">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```