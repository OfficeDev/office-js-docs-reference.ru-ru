
# <a name="userprofile"></a><span data-ttu-id="5e2bb-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="5e2bb-101">userProfile</span></span>

### <span data-ttu-id="5e2bb-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="5e2bb-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e2bb-104">Требования</span><span class="sxs-lookup"><span data-stu-id="5e2bb-104">Requirements</span></span>

|<span data-ttu-id="5e2bb-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="5e2bb-105">Requirement</span></span>| <span data-ttu-id="5e2bb-106">Значение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e2bb-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e2bb-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e2bb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5e2bb-108">1.0</span></span>|
|[<span data-ttu-id="5e2bb-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e2bb-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e2bb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e2bb-110">ReadItem</span></span>|
|[<span data-ttu-id="5e2bb-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e2bb-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e2bb-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="5e2bb-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="5e2bb-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="5e2bb-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-114">displayName :String</span></span>

<span data-ttu-id="5e2bb-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="5e2bb-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5e2bb-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="5e2bb-116">Type:</span></span>

*   <span data-ttu-id="5e2bb-117">String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e2bb-118">Требования</span><span class="sxs-lookup"><span data-stu-id="5e2bb-118">Requirements</span></span>

|<span data-ttu-id="5e2bb-119">Requirement</span><span class="sxs-lookup"><span data-stu-id="5e2bb-119">Requirement</span></span>| <span data-ttu-id="5e2bb-120">Значение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e2bb-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e2bb-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e2bb-122">1.0</span><span class="sxs-lookup"><span data-stu-id="5e2bb-122">1.0</span></span>|
|[<span data-ttu-id="5e2bb-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e2bb-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e2bb-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e2bb-124">ReadItem</span></span>|
|[<span data-ttu-id="5e2bb-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e2bb-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e2bb-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e2bb-127">Пример</span><span class="sxs-lookup"><span data-stu-id="5e2bb-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5e2bb-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-128">emailAddress :String</span></span>

<span data-ttu-id="5e2bb-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="5e2bb-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5e2bb-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="5e2bb-130">Type:</span></span>

*   <span data-ttu-id="5e2bb-131">String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e2bb-132">Требования</span><span class="sxs-lookup"><span data-stu-id="5e2bb-132">Requirements</span></span>

|<span data-ttu-id="5e2bb-133">Requirement</span><span class="sxs-lookup"><span data-stu-id="5e2bb-133">Requirement</span></span>| <span data-ttu-id="5e2bb-134">Значение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e2bb-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e2bb-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e2bb-136">1.0</span><span class="sxs-lookup"><span data-stu-id="5e2bb-136">1.0</span></span>|
|[<span data-ttu-id="5e2bb-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e2bb-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e2bb-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e2bb-138">ReadItem</span></span>|
|[<span data-ttu-id="5e2bb-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e2bb-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e2bb-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e2bb-141">Пример</span><span class="sxs-lookup"><span data-stu-id="5e2bb-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5e2bb-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-142">timeZone :String</span></span>

<span data-ttu-id="5e2bb-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="5e2bb-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5e2bb-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="5e2bb-144">Type:</span></span>

*   <span data-ttu-id="5e2bb-145">String</span><span class="sxs-lookup"><span data-stu-id="5e2bb-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e2bb-146">Требования</span><span class="sxs-lookup"><span data-stu-id="5e2bb-146">Requirements</span></span>

|<span data-ttu-id="5e2bb-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="5e2bb-147">Requirement</span></span>| <span data-ttu-id="5e2bb-148">Значение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e2bb-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e2bb-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e2bb-150">1.0</span><span class="sxs-lookup"><span data-stu-id="5e2bb-150">1.0</span></span>|
|[<span data-ttu-id="5e2bb-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e2bb-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5e2bb-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5e2bb-152">ReadItem</span></span>|
|[<span data-ttu-id="5e2bb-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e2bb-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e2bb-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e2bb-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e2bb-155">Пример</span><span class="sxs-lookup"><span data-stu-id="5e2bb-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```