# <a name="supporturl-element"></a><span data-ttu-id="de5a3-101">Элемент SupportUrl</span><span class="sxs-lookup"><span data-stu-id="de5a3-101">SupportUrl element</span></span>

<span data-ttu-id="de5a3-102">Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.</span><span class="sxs-lookup"><span data-stu-id="de5a3-102">Specifies the URL of a page that provides support information for your add-in.</span></span>

## <a name="syntax"></a><span data-ttu-id="de5a3-103">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="de5a3-103">Syntax</span></span>

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="de5a3-104">Содержащиеся в</span><span class="sxs-lookup"><span data-stu-id="de5a3-104">Contained in</span></span>

[<span data-ttu-id="de5a3-105">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="de5a3-105">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="de5a3-106">Может содержать</span><span class="sxs-lookup"><span data-stu-id="de5a3-106">Can contain</span></span>

|  <span data-ttu-id="de5a3-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="de5a3-107">Element</span></span> | <span data-ttu-id="de5a3-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="de5a3-108">Required</span></span> | <span data-ttu-id="de5a3-109">Описание</span><span class="sxs-lookup"><span data-stu-id="de5a3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="de5a3-110">Переопределение</span><span class="sxs-lookup"><span data-stu-id="de5a3-110">Override</span></span>](override.md)   | <span data-ttu-id="de5a3-111">Нет</span><span class="sxs-lookup"><span data-stu-id="de5a3-111">No</span></span> | <span data-ttu-id="de5a3-112">Задает параметр для URL-адресов дополнительных языковых стандартов</span><span class="sxs-lookup"><span data-stu-id="de5a3-112">Specifies the setting for additional locale urls</span></span> |

## <a name="attributes"></a><span data-ttu-id="de5a3-113">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="de5a3-113">Attributes</span></span>

|<span data-ttu-id="de5a3-114">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="de5a3-114">**Attribute**</span></span>|<span data-ttu-id="de5a3-115">**Тип**</span><span class="sxs-lookup"><span data-stu-id="de5a3-115">**Type**</span></span>|<span data-ttu-id="de5a3-116">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="de5a3-116">**Required**</span></span>|<span data-ttu-id="de5a3-117">**Описание**</span><span class="sxs-lookup"><span data-stu-id="de5a3-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="de5a3-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="de5a3-118">DefaultValue</span></span>|<span data-ttu-id="de5a3-119">URL-адрес</span><span class="sxs-lookup"><span data-stu-id="de5a3-119">URL</span></span>|<span data-ttu-id="de5a3-120">Обязательный</span><span class="sxs-lookup"><span data-stu-id="de5a3-120">required</span></span>|<span data-ttu-id="de5a3-121">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="de5a3-121">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
