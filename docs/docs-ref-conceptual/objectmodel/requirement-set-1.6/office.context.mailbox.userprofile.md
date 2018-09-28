
# <a name="userprofile"></a><span data-ttu-id="2ba46-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="2ba46-101">userProfile</span></span>

### <span data-ttu-id="2ba46-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="2ba46-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ba46-104">Требования</span><span class="sxs-lookup"><span data-stu-id="2ba46-104">Requirements</span></span>

|<span data-ttu-id="2ba46-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="2ba46-105">Requirement</span></span>| <span data-ttu-id="2ba46-106">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ba46-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ba46-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ba46-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2ba46-108">1.0</span></span>|
|[<span data-ttu-id="2ba46-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ba46-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ba46-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ba46-110">ReadItem</span></span>|
|[<span data-ttu-id="2ba46-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ba46-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2ba46-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ba46-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2ba46-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="2ba46-113">Members and methods</span></span>

| <span data-ttu-id="2ba46-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="2ba46-114">Member</span></span> | <span data-ttu-id="2ba46-115">Тип</span><span class="sxs-lookup"><span data-stu-id="2ba46-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2ba46-116">accountType</span><span class="sxs-lookup"><span data-stu-id="2ba46-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="2ba46-117">Member</span><span class="sxs-lookup"><span data-stu-id="2ba46-117">Member</span></span> |
| [<span data-ttu-id="2ba46-118">displayName</span><span class="sxs-lookup"><span data-stu-id="2ba46-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="2ba46-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="2ba46-119">Member</span></span> |
| [<span data-ttu-id="2ba46-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="2ba46-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="2ba46-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="2ba46-121">Member</span></span> |
| [<span data-ttu-id="2ba46-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="2ba46-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="2ba46-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="2ba46-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2ba46-124">Members</span><span class="sxs-lookup"><span data-stu-id="2ba46-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="2ba46-125">accountType: String</span><span class="sxs-lookup"><span data-stu-id="2ba46-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="2ba46-126">Этот член является в настоящее время только поддерживаемые в Outlook 2016 или более поздней версии для Mac (построение 16.9.1212 или более поздней версии).</span><span class="sxs-lookup"><span data-stu-id="2ba46-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="2ba46-127">Получает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="2ba46-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="2ba46-128">В следующей таблице перечислены возможные значения.</span><span class="sxs-lookup"><span data-stu-id="2ba46-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="2ba46-129">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-129">Value</span></span> | <span data-ttu-id="2ba46-130">Описание</span><span class="sxs-lookup"><span data-stu-id="2ba46-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="2ba46-131">Почтовый ящик относится локального сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="2ba46-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="2ba46-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="2ba46-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="2ba46-133">Почтовый ящик связан с Office 365 работы или школе учетной записи.</span><span class="sxs-lookup"><span data-stu-id="2ba46-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="2ba46-134">Почтовый ящик связан с учетной записью личных Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="2ba46-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="2ba46-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="2ba46-135">Type:</span></span>

*   <span data-ttu-id="2ba46-136">String</span><span class="sxs-lookup"><span data-stu-id="2ba46-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ba46-137">Требования</span><span class="sxs-lookup"><span data-stu-id="2ba46-137">Requirements</span></span>

|<span data-ttu-id="2ba46-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="2ba46-138">Requirement</span></span>| <span data-ttu-id="2ba46-139">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ba46-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ba46-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ba46-141">1.6</span><span class="sxs-lookup"><span data-stu-id="2ba46-141">1.6</span></span> |
|[<span data-ttu-id="2ba46-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ba46-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ba46-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ba46-143">ReadItem</span></span>|
|[<span data-ttu-id="2ba46-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ba46-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2ba46-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ba46-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ba46-146">Пример</span><span class="sxs-lookup"><span data-stu-id="2ba46-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="2ba46-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2ba46-147">displayName :String</span></span>

<span data-ttu-id="2ba46-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="2ba46-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2ba46-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="2ba46-149">Type:</span></span>

*   <span data-ttu-id="2ba46-150">String</span><span class="sxs-lookup"><span data-stu-id="2ba46-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ba46-151">Требования</span><span class="sxs-lookup"><span data-stu-id="2ba46-151">Requirements</span></span>

|<span data-ttu-id="2ba46-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="2ba46-152">Requirement</span></span>| <span data-ttu-id="2ba46-153">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ba46-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ba46-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ba46-155">1.0</span><span class="sxs-lookup"><span data-stu-id="2ba46-155">1.0</span></span>|
|[<span data-ttu-id="2ba46-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ba46-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ba46-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ba46-157">ReadItem</span></span>|
|[<span data-ttu-id="2ba46-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ba46-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2ba46-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ba46-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ba46-160">Пример</span><span class="sxs-lookup"><span data-stu-id="2ba46-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2ba46-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2ba46-161">emailAddress :String</span></span>

<span data-ttu-id="2ba46-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="2ba46-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2ba46-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="2ba46-163">Type:</span></span>

*   <span data-ttu-id="2ba46-164">String</span><span class="sxs-lookup"><span data-stu-id="2ba46-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ba46-165">Требования</span><span class="sxs-lookup"><span data-stu-id="2ba46-165">Requirements</span></span>

|<span data-ttu-id="2ba46-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="2ba46-166">Requirement</span></span>| <span data-ttu-id="2ba46-167">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ba46-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ba46-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ba46-169">1.0</span><span class="sxs-lookup"><span data-stu-id="2ba46-169">1.0</span></span>|
|[<span data-ttu-id="2ba46-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ba46-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ba46-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ba46-171">ReadItem</span></span>|
|[<span data-ttu-id="2ba46-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ba46-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2ba46-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ba46-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ba46-174">Пример</span><span class="sxs-lookup"><span data-stu-id="2ba46-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2ba46-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2ba46-175">timeZone :String</span></span>

<span data-ttu-id="2ba46-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2ba46-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2ba46-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="2ba46-177">Type:</span></span>

*   <span data-ttu-id="2ba46-178">String</span><span class="sxs-lookup"><span data-stu-id="2ba46-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ba46-179">Требования</span><span class="sxs-lookup"><span data-stu-id="2ba46-179">Requirements</span></span>

|<span data-ttu-id="2ba46-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="2ba46-180">Requirement</span></span>| <span data-ttu-id="2ba46-181">Значение</span><span class="sxs-lookup"><span data-stu-id="2ba46-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ba46-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ba46-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ba46-183">1.0</span><span class="sxs-lookup"><span data-stu-id="2ba46-183">1.0</span></span>|
|[<span data-ttu-id="2ba46-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ba46-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ba46-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ba46-185">ReadItem</span></span>|
|[<span data-ttu-id="2ba46-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ba46-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2ba46-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ba46-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ba46-188">Пример</span><span class="sxs-lookup"><span data-stu-id="2ba46-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```