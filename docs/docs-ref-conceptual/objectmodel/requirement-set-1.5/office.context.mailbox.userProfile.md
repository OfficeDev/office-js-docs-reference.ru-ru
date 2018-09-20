# <a name="userprofile"></a><span data-ttu-id="651e3-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="651e3-101">userProfile</span></span>

### <span data-ttu-id="651e3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="651e3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="651e3-104">Требования</span><span class="sxs-lookup"><span data-stu-id="651e3-104">Requirements</span></span>

|<span data-ttu-id="651e3-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="651e3-105">Requirement</span></span>| <span data-ttu-id="651e3-106">Значение</span><span class="sxs-lookup"><span data-stu-id="651e3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="651e3-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="651e3-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="651e3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="651e3-108">1.0</span></span>|
|[<span data-ttu-id="651e3-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="651e3-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="651e3-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="651e3-110">ReadItem</span></span>|
|[<span data-ttu-id="651e3-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="651e3-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="651e3-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="651e3-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="651e3-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="651e3-113">Members and methods</span></span>

| <span data-ttu-id="651e3-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="651e3-114">Member</span></span> | <span data-ttu-id="651e3-115">Тип</span><span class="sxs-lookup"><span data-stu-id="651e3-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="651e3-116">displayName</span><span class="sxs-lookup"><span data-stu-id="651e3-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="651e3-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="651e3-117">Member</span></span> |
| [<span data-ttu-id="651e3-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="651e3-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="651e3-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="651e3-119">Member</span></span> |
| [<span data-ttu-id="651e3-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="651e3-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="651e3-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="651e3-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="651e3-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="651e3-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="651e3-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="651e3-123">displayName :String</span></span>

<span data-ttu-id="651e3-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="651e3-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="651e3-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="651e3-125">Type:</span></span>

*   <span data-ttu-id="651e3-126">String</span><span class="sxs-lookup"><span data-stu-id="651e3-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="651e3-127">Требования</span><span class="sxs-lookup"><span data-stu-id="651e3-127">Requirements</span></span>

|<span data-ttu-id="651e3-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="651e3-128">Requirement</span></span>| <span data-ttu-id="651e3-129">Значение</span><span class="sxs-lookup"><span data-stu-id="651e3-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="651e3-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="651e3-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="651e3-131">1.0</span><span class="sxs-lookup"><span data-stu-id="651e3-131">1.0</span></span>|
|[<span data-ttu-id="651e3-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="651e3-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="651e3-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="651e3-133">ReadItem</span></span>|
|[<span data-ttu-id="651e3-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="651e3-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="651e3-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="651e3-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="651e3-136">Пример</span><span class="sxs-lookup"><span data-stu-id="651e3-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="651e3-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="651e3-137">emailAddress :String</span></span>

<span data-ttu-id="651e3-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="651e3-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="651e3-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="651e3-139">Type:</span></span>

*   <span data-ttu-id="651e3-140">String</span><span class="sxs-lookup"><span data-stu-id="651e3-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="651e3-141">Требования</span><span class="sxs-lookup"><span data-stu-id="651e3-141">Requirements</span></span>

|<span data-ttu-id="651e3-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="651e3-142">Requirement</span></span>| <span data-ttu-id="651e3-143">Значение</span><span class="sxs-lookup"><span data-stu-id="651e3-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="651e3-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="651e3-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="651e3-145">1.0</span><span class="sxs-lookup"><span data-stu-id="651e3-145">1.0</span></span>|
|[<span data-ttu-id="651e3-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="651e3-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="651e3-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="651e3-147">ReadItem</span></span>|
|[<span data-ttu-id="651e3-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="651e3-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="651e3-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="651e3-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="651e3-150">Пример</span><span class="sxs-lookup"><span data-stu-id="651e3-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="651e3-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="651e3-151">timeZone :String</span></span>

<span data-ttu-id="651e3-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="651e3-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="651e3-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="651e3-153">Type:</span></span>

*   <span data-ttu-id="651e3-154">String</span><span class="sxs-lookup"><span data-stu-id="651e3-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="651e3-155">Требования</span><span class="sxs-lookup"><span data-stu-id="651e3-155">Requirements</span></span>

|<span data-ttu-id="651e3-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="651e3-156">Requirement</span></span>| <span data-ttu-id="651e3-157">Значение</span><span class="sxs-lookup"><span data-stu-id="651e3-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="651e3-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="651e3-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="651e3-159">1.0</span><span class="sxs-lookup"><span data-stu-id="651e3-159">1.0</span></span>|
|[<span data-ttu-id="651e3-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="651e3-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="651e3-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="651e3-161">ReadItem</span></span>|
|[<span data-ttu-id="651e3-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="651e3-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="651e3-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="651e3-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="651e3-164">Пример</span><span class="sxs-lookup"><span data-stu-id="651e3-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```