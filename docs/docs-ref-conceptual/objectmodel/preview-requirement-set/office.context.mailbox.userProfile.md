
# <a name="userprofile"></a><span data-ttu-id="a8549-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a8549-101">userProfile</span></span>

### <span data-ttu-id="a8549-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a8549-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8549-104">Требования</span><span class="sxs-lookup"><span data-stu-id="a8549-104">Requirements</span></span>

|<span data-ttu-id="a8549-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="a8549-105">Requirement</span></span>| <span data-ttu-id="a8549-106">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8549-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8549-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8549-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a8549-108">1.0</span></span>|
|[<span data-ttu-id="a8549-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8549-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8549-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8549-110">ReadItem</span></span>|
|[<span data-ttu-id="a8549-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8549-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8549-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8549-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a8549-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a8549-113">Members and methods</span></span>

| <span data-ttu-id="a8549-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8549-114">Member</span></span> | <span data-ttu-id="a8549-115">Тип</span><span class="sxs-lookup"><span data-stu-id="a8549-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a8549-116">accountType</span><span class="sxs-lookup"><span data-stu-id="a8549-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a8549-117">Member</span><span class="sxs-lookup"><span data-stu-id="a8549-117">Member</span></span> |
| [<span data-ttu-id="a8549-118">displayName</span><span class="sxs-lookup"><span data-stu-id="a8549-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="a8549-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8549-119">Member</span></span> |
| [<span data-ttu-id="a8549-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a8549-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a8549-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8549-121">Member</span></span> |
| [<span data-ttu-id="a8549-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="a8549-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a8549-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8549-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a8549-124">Members</span><span class="sxs-lookup"><span data-stu-id="a8549-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="a8549-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="a8549-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="a8549-126">Этот член в данный момент только поддерживаемые в Outlook 2016 для Mac, построения 16.9.1212 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="a8549-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="a8549-127">Получает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="a8549-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a8549-128">В следующей таблице перечислены возможные значения.</span><span class="sxs-lookup"><span data-stu-id="a8549-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a8549-129">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-129">Value</span></span> | <span data-ttu-id="a8549-130">Описание</span><span class="sxs-lookup"><span data-stu-id="a8549-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a8549-131">Почтовый ящик относится локального сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="a8549-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a8549-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="a8549-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a8549-133">Почтовый ящик связан с Office 365 работы или школе учетной записи.</span><span class="sxs-lookup"><span data-stu-id="a8549-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a8549-134">Почтовый ящик связан с учетной записью личных Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="a8549-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a8549-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="a8549-135">Type:</span></span>

*   <span data-ttu-id="a8549-136">String</span><span class="sxs-lookup"><span data-stu-id="a8549-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8549-137">Требования</span><span class="sxs-lookup"><span data-stu-id="a8549-137">Requirements</span></span>

|<span data-ttu-id="a8549-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="a8549-138">Requirement</span></span>| <span data-ttu-id="a8549-139">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8549-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8549-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8549-141">1.6</span><span class="sxs-lookup"><span data-stu-id="a8549-141">1.6</span></span> |
|[<span data-ttu-id="a8549-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8549-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8549-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8549-143">ReadItem</span></span>|
|[<span data-ttu-id="a8549-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8549-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8549-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8549-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a8549-146">Пример</span><span class="sxs-lookup"><span data-stu-id="a8549-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="a8549-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a8549-147">displayName :String</span></span>

<span data-ttu-id="a8549-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="a8549-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a8549-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="a8549-149">Type:</span></span>

*   <span data-ttu-id="a8549-150">String</span><span class="sxs-lookup"><span data-stu-id="a8549-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8549-151">Требования</span><span class="sxs-lookup"><span data-stu-id="a8549-151">Requirements</span></span>

|<span data-ttu-id="a8549-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="a8549-152">Requirement</span></span>| <span data-ttu-id="a8549-153">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8549-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8549-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8549-155">1.0</span><span class="sxs-lookup"><span data-stu-id="a8549-155">1.0</span></span>|
|[<span data-ttu-id="a8549-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8549-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8549-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8549-157">ReadItem</span></span>|
|[<span data-ttu-id="a8549-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8549-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8549-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8549-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a8549-160">Пример</span><span class="sxs-lookup"><span data-stu-id="a8549-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a8549-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a8549-161">emailAddress :String</span></span>

<span data-ttu-id="a8549-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="a8549-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a8549-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="a8549-163">Type:</span></span>

*   <span data-ttu-id="a8549-164">String</span><span class="sxs-lookup"><span data-stu-id="a8549-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8549-165">Требования</span><span class="sxs-lookup"><span data-stu-id="a8549-165">Requirements</span></span>

|<span data-ttu-id="a8549-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="a8549-166">Requirement</span></span>| <span data-ttu-id="a8549-167">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8549-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8549-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8549-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a8549-169">1.0</span></span>|
|[<span data-ttu-id="a8549-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8549-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8549-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8549-171">ReadItem</span></span>|
|[<span data-ttu-id="a8549-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8549-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8549-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8549-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a8549-174">Пример</span><span class="sxs-lookup"><span data-stu-id="a8549-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a8549-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a8549-175">timeZone :String</span></span>

<span data-ttu-id="a8549-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a8549-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a8549-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="a8549-177">Type:</span></span>

*   <span data-ttu-id="a8549-178">String</span><span class="sxs-lookup"><span data-stu-id="a8549-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8549-179">Требования</span><span class="sxs-lookup"><span data-stu-id="a8549-179">Requirements</span></span>

|<span data-ttu-id="a8549-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="a8549-180">Requirement</span></span>| <span data-ttu-id="a8549-181">Значение</span><span class="sxs-lookup"><span data-stu-id="a8549-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8549-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8549-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8549-183">1.0</span><span class="sxs-lookup"><span data-stu-id="a8549-183">1.0</span></span>|
|[<span data-ttu-id="a8549-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8549-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8549-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8549-185">ReadItem</span></span>|
|[<span data-ttu-id="a8549-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8549-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8549-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8549-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a8549-188">Пример</span><span class="sxs-lookup"><span data-stu-id="a8549-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```