# <a name="userprofile"></a><span data-ttu-id="b9fc0-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b9fc0-101">userProfile</span></span>

### <span data-ttu-id="b9fc0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b9fc0-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9fc0-104">Требования</span><span class="sxs-lookup"><span data-stu-id="b9fc0-104">Requirements</span></span>

|<span data-ttu-id="b9fc0-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="b9fc0-105">Requirement</span></span>| <span data-ttu-id="b9fc0-106">Значение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9fc0-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9fc0-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9fc0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b9fc0-108">1.0</span></span>|
|[<span data-ttu-id="b9fc0-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b9fc0-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9fc0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9fc0-110">ReadItem</span></span>|
|[<span data-ttu-id="b9fc0-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9fc0-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9fc0-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b9fc0-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="b9fc0-113">Members and methods</span></span>

| <span data-ttu-id="b9fc0-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="b9fc0-114">Member</span></span> | <span data-ttu-id="b9fc0-115">Тип</span><span class="sxs-lookup"><span data-stu-id="b9fc0-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b9fc0-116">displayName</span><span class="sxs-lookup"><span data-stu-id="b9fc0-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="b9fc0-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="b9fc0-117">Member</span></span> |
| [<span data-ttu-id="b9fc0-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b9fc0-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b9fc0-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="b9fc0-119">Member</span></span> |
| [<span data-ttu-id="b9fc0-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="b9fc0-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b9fc0-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="b9fc0-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b9fc0-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="b9fc0-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b9fc0-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-123">displayName :String</span></span>

<span data-ttu-id="b9fc0-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="b9fc0-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b9fc0-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="b9fc0-125">Type:</span></span>

*   <span data-ttu-id="b9fc0-126">String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9fc0-127">Требования</span><span class="sxs-lookup"><span data-stu-id="b9fc0-127">Requirements</span></span>

|<span data-ttu-id="b9fc0-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="b9fc0-128">Requirement</span></span>| <span data-ttu-id="b9fc0-129">Значение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9fc0-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9fc0-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9fc0-131">1.0</span><span class="sxs-lookup"><span data-stu-id="b9fc0-131">1.0</span></span>|
|[<span data-ttu-id="b9fc0-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b9fc0-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9fc0-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9fc0-133">ReadItem</span></span>|
|[<span data-ttu-id="b9fc0-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9fc0-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9fc0-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9fc0-136">Пример</span><span class="sxs-lookup"><span data-stu-id="b9fc0-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b9fc0-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-137">emailAddress :String</span></span>

<span data-ttu-id="b9fc0-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="b9fc0-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b9fc0-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="b9fc0-139">Type:</span></span>

*   <span data-ttu-id="b9fc0-140">String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9fc0-141">Требования</span><span class="sxs-lookup"><span data-stu-id="b9fc0-141">Requirements</span></span>

|<span data-ttu-id="b9fc0-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="b9fc0-142">Requirement</span></span>| <span data-ttu-id="b9fc0-143">Значение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9fc0-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9fc0-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9fc0-145">1.0</span><span class="sxs-lookup"><span data-stu-id="b9fc0-145">1.0</span></span>|
|[<span data-ttu-id="b9fc0-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b9fc0-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9fc0-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9fc0-147">ReadItem</span></span>|
|[<span data-ttu-id="b9fc0-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9fc0-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9fc0-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9fc0-150">Пример</span><span class="sxs-lookup"><span data-stu-id="b9fc0-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b9fc0-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-151">timeZone :String</span></span>

<span data-ttu-id="b9fc0-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="b9fc0-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b9fc0-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="b9fc0-153">Type:</span></span>

*   <span data-ttu-id="b9fc0-154">String</span><span class="sxs-lookup"><span data-stu-id="b9fc0-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9fc0-155">Требования</span><span class="sxs-lookup"><span data-stu-id="b9fc0-155">Requirements</span></span>

|<span data-ttu-id="b9fc0-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="b9fc0-156">Requirement</span></span>| <span data-ttu-id="b9fc0-157">Значение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9fc0-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9fc0-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9fc0-159">1.0</span><span class="sxs-lookup"><span data-stu-id="b9fc0-159">1.0</span></span>|
|[<span data-ttu-id="b9fc0-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b9fc0-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9fc0-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b9fc0-161">ReadItem</span></span>|
|[<span data-ttu-id="b9fc0-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9fc0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b9fc0-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9fc0-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9fc0-164">Пример</span><span class="sxs-lookup"><span data-stu-id="b9fc0-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```