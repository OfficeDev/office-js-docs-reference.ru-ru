# <a name="sourcelocation-element"></a><span data-ttu-id="03887-101">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="03887-101">SourceLocation element</span></span>

<span data-ttu-id="03887-p101">Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="03887-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="03887-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="03887-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="03887-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="03887-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="03887-106">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="03887-106">Contained in</span></span>

- <span data-ttu-id="03887-107">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="03887-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="03887-108">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="03887-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="03887-109">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="03887-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="03887-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="03887-110">Can contain</span></span>

[<span data-ttu-id="03887-111">Переопределение</span><span class="sxs-lookup"><span data-stu-id="03887-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="03887-112">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="03887-112">Attributes</span></span>

|<span data-ttu-id="03887-113">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="03887-113">**Attribute**</span></span>|<span data-ttu-id="03887-114">**Тип**</span><span class="sxs-lookup"><span data-stu-id="03887-114">**Type**</span></span>|<span data-ttu-id="03887-115">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="03887-115">**Required**</span></span>|<span data-ttu-id="03887-116">**Описание**</span><span class="sxs-lookup"><span data-stu-id="03887-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="03887-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="03887-117">DefaultValue</span></span>|<span data-ttu-id="03887-118">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="03887-118">URL</span></span>|<span data-ttu-id="03887-119">Обязательный</span><span class="sxs-lookup"><span data-stu-id="03887-119">required</span></span>|<span data-ttu-id="03887-120">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="03887-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
